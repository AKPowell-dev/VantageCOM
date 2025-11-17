using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class EditingService
    {
        private readonly Excel.Application _app;

        public EditingService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public void AdjustNumbers(int count, bool subtract, bool grow)
        {
            if (count <= 0)
            {
                count = 1;
            }

            if (!RangeHelpers.TryGetActiveRange(_app, out var selection))
            {
                return;
            }

            double totalCells = Convert.ToDouble(selection.Cells.CountLarge);
            if (totalCells > 1048576d * 8d)
            {
                ShowStatusTemporarily("Too many cells selected.");
                return;
            }

            if (selection.Count == 1)
            {
                var formula = Convert.ToString(selection.Formula);
                if (string.IsNullOrEmpty(formula))
                {
                    selection.Value2 = subtract ? -count : count;
                    return;
                }
            }

            var prevCalc = _app.Calculation;
            var prevCursor = _app.Cursor;
            UiGuard guard = null;

            try
            {
                bool multiCell = totalCells > 1;
                if (multiCell)
                {
                    guard = new UiGuard(_app);
                    _app.Calculation = Excel.XlCalculation.xlCalculationManual;
                }

                _app.Cursor = Excel.XlMousePointer.xlWait;

                int procSign = subtract ? -1 : 1;
                double inc = count;
                int rowInc = grow && selection.Rows.Count > 1 ? procSign : 0;
                int colInc = grow && selection.Rows.Count < 2 ? procSign : 0;
                long processed = 0;
                long total = Convert.ToInt64(totalCells);
                double startTime = Environment.TickCount / 1000.0;

                foreach (Excel.Range cell in selection.Cells)
                {
                    try
                    {
                        if (!TryProcessCell(cell, inc * procSign, procSign, count))
                        {
                            // nothing to do
                        }
                    }
                    catch
                    {
                        // ignore cell errors
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cell);
                    }

                    inc += count * procSign * colInc;
                    processed++;

                    if ((processed & 0xFFF) == 0)
                    {
                        UpdateProgress(processed, total, ref startTime);
                    }
                }

                if (selection.Rows.Count > 1)
                {
                    inc += count * procSign * rowInc;
                }
            }
            finally
            {
                _app.StatusBar = false;
                _app.Cursor = prevCursor;
                _app.Calculation = prevCalc;
                guard?.Dispose();
            }
        }

        public void ApplyAutoFill()
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var selection))
            {
                return;
            }

            Excel.Range baseRange = null;
            try
            {
                baseRange = DetermineBaseRange(selection);
                if (!RangeHelpers.IsRangeValid(baseRange))
                {
                    return;
                }

                bool singleNumeric = baseRange.Count == 1 && double.TryParse(Convert.ToString(baseRange.Value2), out _);
                if (singleNumeric)
                {
                    baseRange.AutoFill(selection, Excel.XlAutoFillType.xlFillSeries);
                }
                else
                {
                    baseRange.AutoFill(selection, Excel.XlAutoFillType.xlFillDefault);
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(baseRange);
            }
        }

        public void NavigateSpecialCells(int typeValue, Excel.XlSearchOrder searchOrder, bool forward, int steps)
        {
            if (steps <= 0)
            {
                steps = 1;
            }

            var sheet = _app.ActiveSheet as Excel.Worksheet;
            if (sheet == null)
            {
                return;
            }

            Excel.Range found = null;
            Excel.Range current = null;
            try
            {
                var used = sheet.UsedRange;
                found = used.SpecialCells((Excel.XlCellType)typeValue);
                current = _app.ActiveCell;
                if (found == null || current == null)
                {
                    return;
                }

                Excel.XlSearchDirection direction = forward ? Excel.XlSearchDirection.xlNext : Excel.XlSearchDirection.xlPrevious;
                int cycle = Math.Max(steps, 1);
                Excel.Range target = current;
                for (int i = 0; i < cycle; i++)
                {
                    target = DetermineSpecialCell(target, found, (Excel.XlCellType)typeValue, searchOrder, direction);
                    if (target == null)
                    {
                        break;
                    }
                }

                if (RangeHelpers.IsRangeValid(target))
                {
                    RangeHelpers.SafeSelect(target);
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(found);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(current);
            }
        }

        public void SubstituteType(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            string payload = text.Length == 1 ? text : text.Substring(0, 1);
            try
            {
                SendKeys.SendWait(payload);
            }
            catch
            {
                // ignore send failures
            }
        }

        public void InsertRows(int count, bool append)
        {
            if (count < 1) count = 1;
            using (new UiGuard(_app))
            {
                var target = ResolveRowRange(RowTargetType.Entire, count);
                if (!RangeHelpers.IsRangeValid(target))
                {
                    return;
                }

                try
                {
                    if (append)
                    {
                        int lastRow = target.Row + target.Rows.Count - 1;
                        int maxRow = Convert.ToInt32(target.Worksheet.Rows.Count);
                        if (lastRow < maxRow)
                        {
                            target = target.Offset[1, 0];
                        }
                    }

                    target.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseCom(target);
                }
            }
        }

        public void DeleteRows(int rawTargetType, int count)
        {
            if (count < 1) count = 1;
            var targetType = NormalizeRowTarget(rawTargetType);
            using (new UiGuard(_app))
            {
                var target = ResolveRowRange(targetType, count);
                if (!RangeHelpers.IsRangeValid(target))
                {
                    return;
                }

                try
                {
                    target.EntireRow.Delete();
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseCom(target);
                }
            }
        }

        public void HideRows(int rawTargetType, int count, bool hide)
        {
            if (count < 1) count = 1;
            var targetType = NormalizeRowTarget(rawTargetType);
            using (new UiGuard(_app))
            {
                var target = ResolveRowRange(targetType, count);
                if (!RangeHelpers.IsRangeValid(target))
                {
                    return;
                }

                try
                {
                    target.EntireRow.Hidden = hide;
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseCom(target);
                }
            }
        }

        public void GroupRows(int count, bool group)
        {
            if (count < 1) count = 1;
            using (new UiGuard(_app))
            {
                var target = ResolveRowRange(RowTargetType.Entire, count);
                if (!RangeHelpers.IsRangeValid(target))
                {
                    return;
                }

                try
                {
                    if (group)
                    {
                        target.Rows.Group();
                    }
                    else
                    {
                        target.Rows.Ungroup();
                    }
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseCom(target);
                }
            }
        }

        public void InsertColumns(int count, bool append)
        {
            if (count < 1) count = 1;
            using (new UiGuard(_app))
            {
                var target = ResolveColumnRange(ColumnTargetType.Entire, count);
                if (!RangeHelpers.IsRangeValid(target))
                {
                    return;
                }

                try
                {
                    if (append)
                    {
                        int lastCol = target.Column + target.Columns.Count - 1;
                        int maxCol = Convert.ToInt32(target.Worksheet.Columns.Count);
                        if (lastCol < maxCol)
                        {
                            target = target.Offset[0, 1];
                        }
                    }

                    target.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseCom(target);
                }
            }
        }

        public void DeleteColumns(int rawTargetType, int count)
        {
            if (count < 1) count = 1;
            var targetType = NormalizeColumnTarget(rawTargetType);
            using (new UiGuard(_app))
            {
                var target = ResolveColumnRange(targetType, count);
                if (!RangeHelpers.IsRangeValid(target))
                {
                    return;
                }

                try
                {
                    target.EntireColumn.Delete();
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseCom(target);
                }
            }
        }

        public void HideColumns(int rawTargetType, int count, bool hide)
        {
            if (count < 1) count = 1;
            var targetType = NormalizeColumnTarget(rawTargetType);
            using (new UiGuard(_app))
            {
                var target = ResolveColumnRange(targetType, count);
                if (!RangeHelpers.IsRangeValid(target))
                {
                    return;
                }

                try
                {
                    target.EntireColumn.Hidden = hide;
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseCom(target);
                }
            }
        }

        public void GroupColumns(int count, bool group)
        {
            if (count < 1) count = 1;
            using (new UiGuard(_app))
            {
                var target = ResolveColumnRange(ColumnTargetType.Entire, count);
                if (!RangeHelpers.IsRangeValid(target))
                {
                    return;
                }

                try
                {
                    if (group)
                    {
                        target.Columns.Group();
                    }
                    else
                    {
                        target.Columns.Ungroup();
                    }
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseCom(target);
                }
            }
        }

        private bool TryProcessCell(Excel.Range cell, double delta, int procSign, int step)
        {
            string formula = Convert.ToString(cell.Formula);
            if (string.IsNullOrEmpty(formula))
            {
                return false;
            }

            if (formula.IndexOf('=') >= 0)
            {
                return false;
            }

            object valueObj = cell.Value2;
            if (valueObj == null)
            {
                return false;
            }

            if (valueObj is double || valueObj is float || valueObj is decimal || valueObj is int || valueObj is short || valueObj is long)
            {
                double current = Convert.ToDouble(valueObj);
                string numberFormat = Convert.ToString(cell.NumberFormatLocal) ?? string.Empty;
                bool isPercent = numberFormat.IndexOf("%", StringComparison.Ordinal) >= 0;
                double newValue = isPercent ? current + (delta / 100d) : current + delta;
                cell.Value2 = newValue;
                return true;
            }

            string text = Convert.ToString(valueObj);
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            string updated = TryAdjustTextNumber(text, delta);
            if (updated == null)
            {
                return false;
            }

            bool keepPrefix = false;
            try
            {
                keepPrefix = Convert.ToString(cell.PrefixCharacter) == "'";
            }
            catch
            {
                keepPrefix = false;
            }

            cell.Value2 = keepPrefix ? "'" + updated : updated;
            return true;
        }

        private static string TryAdjustTextNumber(string text, double delta)
        {
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            bool onlyNumeric = text.All(ch => char.IsDigit(ch) || ch == '.' || ch == '-' || ch == '+');
            if (onlyNumeric && double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var numeric))
            {
                return (numeric + delta).ToString(CultureInfo.InvariantCulture);
            }

            int length = text.Length;
            if (char.IsDigit(text[length - 1]))
            {
                int len = Math.Min(10, length);
                int digits = 1;
                while (digits < len && char.IsDigit(text[length - digits - 1]))
                {
                    digits++;
                }

                string prefix = text.Substring(0, length - digits);
                string suffixDigits = text.Substring(length - digits);
                if (double.TryParse(suffixDigits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var numericSuffix))
                {
                    double result = Math.Max(numericSuffix + delta, 0);
                    string formatted = Convert.ToInt64(Math.Round(result, MidpointRounding.AwayFromZero)).ToString(new string('0', digits), CultureInfo.InvariantCulture);
                    return prefix + formatted;
                }
            }

            if (char.IsDigit(text[0]))
            {
                int len = Math.Min(10, length);
                int digits = 1;
                while (digits < len && char.IsDigit(text[digits]))
                {
                    digits++;
                }

                string digitPart = text.Substring(0, digits);
                if (double.TryParse(digitPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out var numericPrefix))
                {
                    double result = Math.Max(numericPrefix + delta, 0);
                    string formatted = Convert.ToInt64(Math.Round(result, MidpointRounding.AwayFromZero)).ToString(new string('0', digits), CultureInfo.InvariantCulture);
                    return formatted + text.Substring(digits);
                }
            }

            return null;
        }

        private void UpdateProgress(long processed, long total, ref double startTime)
        {
            if (total <= 0)
            {
                return;
            }

            double currentTime = Environment.TickCount / 1000.0;
            if (currentTime - startTime > 0.5d)
            {
                _app.StatusBar = $"Processing numbers... {processed} / {total}";
                startTime = currentTime;
                Application.DoEvents();
            }
        }

        private void ShowStatusTemporarily(string message)
        {
            try
            {
                _app.StatusBar = message;
                Application.DoEvents();
            }
            finally
            {
                _app.StatusBar = false;
            }
        }

        private Excel.Range DetermineBaseRange(Excel.Range selection)
        {
            if (selection == null)
            {
                return null;
            }

            if (selection.Columns.Count > 1 && selection.Rows.Count > 1)
            {
                return DetermineBaseRangeMatrix(selection);
            }

            return DetermineBaseRangeLine(selection);
        }

        private Excel.Range DetermineBaseRangeMatrix(Excel.Range selection)
        {
            Excel.Range result = null;
            try
            {
                double avgTop = CountA(selection.Worksheet.Range[selection.Cells[1, 1], selection.Cells[1, selection.Columns.Count]]) / (double)selection.Columns.Count;
                double avgLeft = CountA(selection.Worksheet.Range[selection.Cells[1, 1], selection.Cells[selection.Rows.Count, 1]]) / (double)selection.Rows.Count;
                double avgBottom = CountA(selection.Worksheet.Range[selection.Cells[selection.Rows.Count, 1], selection.Cells[selection.Rows.Count, selection.Columns.Count]]) / (double)selection.Columns.Count;
                double avgRight = CountA(selection.Worksheet.Range[selection.Cells[1, selection.Columns.Count], selection.Cells[selection.Rows.Count, selection.Columns.Count]]) / (double)selection.Rows.Count;

                var topLeft = selection.Cells[1, 1] as Excel.Range;
                var bottomRight = selection.Cells[selection.Rows.Count, selection.Columns.Count] as Excel.Range;

                if (string.IsNullOrEmpty(Convert.ToString(topLeft.Formula)))
                {
                    avgTop = 0;
                    avgLeft = 0;
                }

                if (string.IsNullOrEmpty(Convert.ToString(bottomRight.Formula)))
                {
                    avgBottom = 0;
                    avgRight = 0;
                }

                double avgMax = new[] { avgTop, avgLeft, avgBottom, avgRight }.Max();
                var ws = selection.Worksheet;

                if (Math.Abs(avgMax - avgTop) < double.Epsilon)
                {
                    var start = selection.Cells[1, 1];
                    var end = ws.Range[selection.Cells[1, 1], selection.Cells[1, selection.Columns.Count]];
                    result = ws.Range[start, end];
                    var extension = InnerDataSearch(result, Direction.TopToBottom, selection.Rows.Count - 1);
                    result = ws.Range[result, extension];
                }
                else if (Math.Abs(avgMax - avgLeft) < double.Epsilon)
                {
                    var start = selection.Cells[1, 1];
                    var end = ws.Range[selection.Cells[selection.Rows.Count, 1], selection.Cells[selection.Rows.Count, 1]];
                    result = ws.Range[start, end];
                    var extension = InnerDataSearch(result, Direction.LeftToRight, selection.Columns.Count - 1);
                    result = ws.Range[result, extension];
                }
                else if (Math.Abs(avgMax - avgBottom) < double.Epsilon)
                {
                    var start = ws.Range[selection.Cells[selection.Rows.Count, 1], selection.Cells[selection.Rows.Count, selection.Columns.Count]];
                    var extension = InnerDataSearch(start, Direction.BottomToTop, selection.Rows.Count - 1);
                    result = ws.Range[extension, start];
                }
                else
                {
                    var start = ws.Range[selection.Cells[1, selection.Columns.Count], selection.Cells[selection.Rows.Count, selection.Columns.Count]];
                    var extension = InnerDataSearch(start, Direction.RightToLeft, selection.Columns.Count - 1);
                    result = ws.Range[extension, start];
                }

                return result;
            }
            catch
            {
                return null;
            }
        }

        private Excel.Range DetermineBaseRangeLine(Excel.Range selection)
        {
            try
            {
                var first = selection.Cells[1, 1] as Excel.Range;
                var last = selection.Cells[selection.Cells.Count] as Excel.Range;
                var ws = selection.Worksheet;

                if (!string.IsNullOrEmpty(Convert.ToString(first?.Formula)))
                {
                    if (selection.Cells.Count > 1)
                    {
                        var second = selection.Cells[2] as Excel.Range;
                        if (second != null && !string.IsNullOrEmpty(Convert.ToString(second.Formula)))
                        {
                            if (selection.Columns.Count > 1)
                            {
                                var end = first.End[Excel.XlDirection.xlToRight];
                                return ws.Range[first, end];
                            }
                            else
                            {
                                var end = first.End[Excel.XlDirection.xlDown];
                                return ws.Range[first, end];
                            }
                        }
                    }

                    return first;
                }

                if (!string.IsNullOrEmpty(Convert.ToString(last?.Formula)))
                {
                    if (selection.Cells.Count > 1)
                    {
                        var previous = selection.Cells[selection.Cells.Count - 1] as Excel.Range;
                        if (previous != null && !string.IsNullOrEmpty(Convert.ToString(previous.Formula)))
                        {
                            if (selection.Columns.Count > 1)
                            {
                                var start = last.End[Excel.XlDirection.xlToLeft];
                                return ws.Range[start, last];
                            }
                            else
                            {
                                var start = last.End[Excel.XlDirection.xlUp];
                                return ws.Range[start, last];
                            }
                        }
                    }

                    return last;
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private Excel.Range InnerDataSearch(Excel.Range target, Direction direction, int limit, int searchCount = 0, double expected = 0)
        {
            if (target == null || limit < 0)
            {
                return target;
            }

            if (searchCount > limit)
            {
                return target;
            }

            double nonBlank = CountA(target);
            if (searchCount == 0 || Math.Abs(nonBlank - expected) < double.Epsilon)
            {
                var next = OffsetRange(target, direction);
                var recursive = InnerDataSearch(next, direction, limit, searchCount + 1, nonBlank);
                if (RangeHelpers.IsRangeValid(recursive))
                {
                    return recursive;
                }
            }

            return target;
        }

        private Excel.Range OffsetRange(Excel.Range range, Direction direction)
        {
            if (!RangeHelpers.IsRangeValid(range))
            {
                return range;
            }

            try
            {
                int rowOffset = 0;
                int colOffset = 0;
                switch (direction)
                {
                    case Direction.TopToBottom:
                        rowOffset = 1;
                        break;
                    case Direction.BottomToTop:
                        rowOffset = -1;
                        break;
                    case Direction.LeftToRight:
                        colOffset = 1;
                        break;
                    case Direction.RightToLeft:
                        colOffset = -1;
                        break;
                }

                return range.Offset[rowOffset, colOffset];
            }
            catch
            {
                return range;
            }
        }

        private Excel.Range DetermineSpecialCell(Excel.Range baseCell, Excel.Range pool, Excel.XlCellType type, Excel.XlSearchOrder order, Excel.XlSearchDirection direction)
        {
            try
            {
                if (baseCell == null || pool == null)
                {
                    return null;
                }

                var sheet = baseCell.Worksheet;
                if (sheet == null)
                {
                    return null;
                }

                Excel.Range used = sheet.UsedRange;
                int minRow = used.Row;
                int minCol = used.Column;
                int maxRow = used.Row + used.Rows.Count - 1;
                int maxCol = used.Column + used.Columns.Count - 1;

                Excel.Range check = sheet.Range[sheet.Cells[minRow, minCol], sheet.Cells[maxRow, maxCol]];
                Excel.Range intersection = null;
                try
                {
                    intersection = _app.Intersect(check, pool);
                    if (intersection == null)
                    {
                        return null;
                    }

                    return ClosestSearch(intersection, order, direction, type == Excel.XlCellType.xlCellTypeBlanks);
                }
                finally
                {
                    ReleaseCom(check);
                    ReleaseCom(used);
                    ReleaseCom(intersection);
                }
            }
            catch
            {
                return null;
            }
        }

        private Excel.Range ClosestSearch(Excel.Range candidates, Excel.XlSearchOrder order, Excel.XlSearchDirection direction, bool checkMergedBlank)
        {
            if (candidates == null)
            {
                return null;
            }

            Excel.Range best = null;
            var areas = candidates.Areas as Excel.Areas;
            if (areas == null)
            {
                return candidates;
            }

            for (int idx = 1; idx <= areas.Count; idx++)
            {
                Excel.Range area = null;
                try
                {
                    area = areas[idx] as Excel.Range;
                    if (area == null)
                    {
                        continue;
                    }

                    Excel.Range candidate = direction == Excel.XlSearchDirection.xlNext
                        ? area.Cells[1, 1] as Excel.Range
                        : area.Cells[area.Rows.Count, area.Columns.Count] as Excel.Range;
                    if (candidate == null)
                    {
                        continue;
                    }

                    if (Convert.ToBoolean(candidate.MergeCells))
                    {
                        var mergeArea = candidate.MergeArea as Excel.Range;
                        var mergeFirst = mergeArea?.Cells[1, 1] as Excel.Range;
                        if (mergeFirst == null)
                        {
                            candidate = null;
                        }
                        else if (checkMergedBlank && !string.IsNullOrEmpty(Convert.ToString(mergeFirst.Value2)))
                        {
                            candidate = null;
                        }
                        else
                        {
                            candidate = mergeFirst;
                        }
                    }

                    if (candidate == null)
                    {
                        continue;
                    }

                    if (best == null)
                    {
                        best = candidate;
                        continue;
                    }

                    bool replace = false;
                    if (order == Excel.XlSearchOrder.xlByColumns)
                    {
                        if (direction == Excel.XlSearchDirection.xlNext)
                        {
                            replace = candidate.Column < best.Column ||
                                      (candidate.Column == best.Column && candidate.Row < best.Row);
                        }
                        else
                        {
                            replace = candidate.Column > best.Column ||
                                      (candidate.Column == best.Column && candidate.Row > best.Row);
                        }
                    }
                    else
                    {
                        if (direction == Excel.XlSearchDirection.xlNext)
                        {
                            replace = candidate.Row < best.Row ||
                                      (candidate.Row == best.Row && candidate.Column < best.Column);
                        }
                        else
                        {
                            replace = candidate.Row > best.Row ||
                                      (candidate.Row == best.Row && candidate.Column > best.Column);
                        }
                    }

                    if (replace)
                    {
                        best = candidate;
                    }
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseCom(area);
                }
            }
            ReleaseCom(areas);

            return best;
        }

        private Excel.Range ResolveRowRange(RowTargetType type, int count)
        {
            Excel.Range selection = null;
            try { selection = _app.Selection as Excel.Range; } catch { }
            Excel.Worksheet sheet = selection?.Worksheet ?? _app.ActiveSheet as Excel.Worksheet;
            if (sheet == null)
            {
                return null;
            }

            Excel.Range activeCell = null;
            try { activeCell = _app.ActiveCell as Excel.Range; } catch { }

            int maxRow = Convert.ToInt32(sheet.Rows.Count);
            count = Math.Max(1, count);

            int startRow = 1, endRow = 1;
            bool hasRange = true;

            switch (type)
            {
                case RowTargetType.Entire:
                    if (selection != null)
                    {
                        if (selection.Rows.Count > 1 || count == 1)
                        {
                            return selection.EntireRow;
                        }
                        startRow = selection.Row;
                    }
                    else if (activeCell != null)
                    {
                        startRow = activeCell.Row;
                    }
                    endRow = Math.Min(maxRow, startRow + count - 1);
                    break;
                case RowTargetType.ToFirstRows:
                    startRow = 1;
                    endRow = Math.Max(1, activeCell?.Row ?? 1);
                    break;
                case RowTargetType.ToTopRows:
                    GetUsedBounds(sheet, out var usedTop, out _, out _, out _);
                    startRow = usedTop;
                    endRow = Math.Max(usedTop, activeCell?.Row ?? usedTop);
                    if (startRow > endRow) hasRange = false;
                    break;
                case RowTargetType.ToBottomRows:
                    GetUsedBounds(sheet, out _, out var usedBottom, out _, out _);
                    startRow = activeCell?.Row ?? 1;
                    endRow = Math.Max(startRow, usedBottom);
                    if (startRow > endRow) hasRange = false;
                    break;
                case RowTargetType.ToTopOfCurrentRegionRows:
                    if (activeCell == null)
                    {
                        hasRange = false;
                        break;
                    }
                    TryGetCurrentRegionBounds(activeCell, out startRow, out _, out _, out _);
                    endRow = activeCell.Row;
                    if (startRow > endRow) hasRange = false;
                    break;
                case RowTargetType.ToBottomOfCurrentRegionRows:
                    if (activeCell == null)
                    {
                        hasRange = false;
                        break;
                    }
                    TryGetCurrentRegionBounds(activeCell, out _, out endRow, out _, out _);
                    startRow = activeCell.Row;
                    if (startRow > endRow) hasRange = false;
                    break;
                case RowTargetType.UsedRangeRows:
                    var used = sheet.UsedRange;
                    if (used == null)
                    {
                        return null;
                    }
                    ReleaseCom(used);
                    return sheet.UsedRange.EntireRow;
                default:
                    return selection?.EntireRow;
            }

            if (!hasRange)
            {
                return null;
            }

            startRow = Math.Max(1, Math.Min(startRow, maxRow));
            endRow = Math.Max(1, Math.Min(endRow, maxRow));
            if (startRow > endRow)
            {
                return null;
            }

            return sheet.Range[sheet.Rows[startRow], sheet.Rows[endRow]].EntireRow;
        }

        private Excel.Range ResolveColumnRange(ColumnTargetType type, int count)
        {
            Excel.Range selection = null;
            try { selection = _app.Selection as Excel.Range; } catch { }
            Excel.Worksheet sheet = selection?.Worksheet ?? _app.ActiveSheet as Excel.Worksheet;
            if (sheet == null)
            {
                return null;
            }

            Excel.Range activeCell = null;
            try { activeCell = _app.ActiveCell as Excel.Range; } catch { }

            int maxCol = Convert.ToInt32(sheet.Columns.Count);
            count = Math.Max(1, count);

            int startCol = 1, endCol = 1;
            bool hasRange = true;

            switch (type)
            {
                case ColumnTargetType.Entire:
                    if (selection != null)
                    {
                        if (selection.Columns.Count > 1 || count == 1)
                        {
                            return selection.EntireColumn;
                        }
                        startCol = selection.Column;
                    }
                    else if (activeCell != null)
                    {
                        startCol = activeCell.Column;
                    }
                    endCol = Math.Min(maxCol, startCol + count - 1);
                    break;
                case ColumnTargetType.ToFirstColumns:
                    startCol = 1;
                    endCol = Math.Max(1, activeCell?.Column ?? 1);
                    break;
                case ColumnTargetType.ToLeftEndColumns:
                    GetUsedBounds(sheet, out _, out _, out var usedLeft, out _);
                    startCol = usedLeft;
                    endCol = Math.Max(usedLeft, activeCell?.Column ?? usedLeft);
                    if (startCol > endCol) hasRange = false;
                    break;
                case ColumnTargetType.ToRightEndColumns:
                    GetUsedBounds(sheet, out _, out _, out _, out var usedRight);
                    startCol = activeCell?.Column ?? 1;
                    endCol = Math.Max(startCol, usedRight);
                    if (startCol > endCol) hasRange = false;
                    break;
                case ColumnTargetType.ToLeftOfCurrentRegionColumns:
                    if (activeCell == null)
                    {
                        hasRange = false;
                        break;
                    }
                    TryGetCurrentRegionBounds(activeCell, out _, out _, out startCol, out _);
                    endCol = activeCell.Column;
                    if (startCol > endCol) hasRange = false;
                    break;
                case ColumnTargetType.ToRightOfCurrentRegionColumns:
                    if (activeCell == null)
                    {
                        hasRange = false;
                        break;
                    }
                    TryGetCurrentRegionBounds(activeCell, out _, out _, out _, out endCol);
                    startCol = activeCell.Column;
                    if (startCol > endCol) hasRange = false;
                    break;
                case ColumnTargetType.UsedRangeColumns:
                    var usedRange = sheet.UsedRange;
                    if (usedRange == null)
                    {
                        return null;
                    }
                    ReleaseCom(usedRange);
                    return sheet.UsedRange.EntireColumn;
                default:
                    return selection?.EntireColumn;
            }

            if (!hasRange)
            {
                return null;
            }

            startCol = Math.Max(1, Math.Min(startCol, maxCol));
            endCol = Math.Max(1, Math.Min(endCol, maxCol));
            if (startCol > endCol)
            {
                return null;
            }

            return sheet.Range[sheet.Columns[startCol], sheet.Columns[endCol]].EntireColumn;
        }

        private RowTargetType NormalizeRowTarget(int value)
        {
            if (value < 0 || value > (int)RowTargetType.UsedRangeRows)
            {
                return RowTargetType.Entire;
            }

            return (RowTargetType)value;
        }

        private ColumnTargetType NormalizeColumnTarget(int value)
        {
            if (value < 0 || value > (int)ColumnTargetType.UsedRangeColumns)
            {
                return ColumnTargetType.Entire;
            }

            return (ColumnTargetType)value;
        }

        private void GetUsedBounds(Excel.Worksheet sheet, out int firstRow, out int lastRow, out int firstCol, out int lastCol)
        {
            firstRow = 1;
            firstCol = 1;
            lastRow = Convert.ToInt32(sheet.Rows.Count);
            lastCol = Convert.ToInt32(sheet.Columns.Count);

            try
            {
                var used = sheet.UsedRange;
                if (used != null)
                {
                    firstRow = used.Row;
                    firstCol = used.Column;
                    lastRow = firstRow + used.Rows.Count - 1;
                    lastCol = firstCol + used.Columns.Count - 1;
                    ReleaseCom(used);
                }
            }
            catch
            {
                // ignore
            }

            firstRow = Math.Max(1, firstRow);
            firstCol = Math.Max(1, firstCol);
        }

        private void TryGetCurrentRegionBounds(Excel.Range cell, out int firstRow, out int lastRow, out int firstCol, out int lastCol)
        {
            firstRow = cell.Row;
            lastRow = cell.Row;
            firstCol = cell.Column;
            lastCol = cell.Column;

            try
            {
                var region = cell.CurrentRegion;
                if (region != null)
                {
                    firstRow = region.Row;
                    firstCol = region.Column;
                    lastRow = firstRow + region.Rows.Count - 1;
                    lastCol = firstCol + region.Columns.Count - 1;
                    ReleaseCom(region);
                }
            }
            catch
            {
                // ignore
            }
        }

        private double CountA(Excel.Range range)
        {
            try
            {
                return Convert.ToDouble(_app.WorksheetFunction.CountA(range));
            }
            catch
            {
                return 0;
            }
        }

        private void ReleaseCom(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch
            {
                // ignore
            }
        }

        private enum Direction
        {
            TopToBottom,
            LeftToRight,
            BottomToTop,
            RightToLeft
        }

        private enum RowTargetType
        {
            Entire = 0,
            ToTopRows = 1,
            ToBottomRows = 2,
            ToTopOfCurrentRegionRows = 3,
            ToBottomOfCurrentRegionRows = 4,
            ToFirstRows = 5,
            UsedRangeRows = 6
        }

        private enum ColumnTargetType
        {
            Entire = 0,
            ToLeftEndColumns = 1,
            ToRightEndColumns = 2,
            ToLeftOfCurrentRegionColumns = 3,
            ToRightOfCurrentRegionColumns = 4,
            ToFirstColumns = 5,
            UsedRangeColumns = 6
        }
    }
}
