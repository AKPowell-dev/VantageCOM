using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
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

                double singleNumericValue;
                bool singleNumeric = baseRange.Count == 1 && double.TryParse(Convert.ToString(baseRange.Value2), out singleNumericValue);
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

        public void AdjustRowHeight(int delta)
        {
            if (delta == 0)
            {
                return;
            }

            if (!TryGetRowTarget(out var targetRows, out var currentHeight))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                var updated = Math.Max(0.0, Math.Min(409.5, currentHeight + delta));
                try { targetRows.RowHeight = updated; }
                catch { }
                finally { ReleaseCom(targetRows); }
            }
        }

        public void AdjustColumnWidth(int delta)
        {
            if (delta == 0)
            {
                return;
            }

            if (!TryGetColumnTarget(out var targetColumns, out var currentWidth))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                var updated = Math.Max(0.0, Math.Min(255.0, currentWidth + delta));
                try { targetColumns.ColumnWidth = updated; }
                catch { }
                finally { ReleaseCom(targetColumns); }
            }
        }

        public void MoveActiveCellBy(int rowDelta, int columnDelta)
        {
            if (rowDelta == 0 && columnDelta == 0)
            {
                return;
            }

            Excel.Range activeCell = null;
            Excel.Worksheet sheet = null;
            Excel.Range target = null;
            var modifiers = Control.ModifierKeys;
            bool ctrlPressed = (modifiers & Keys.Control) == Keys.Control;
            bool shiftPressed = (modifiers & Keys.Shift) == Keys.Shift;

            try
            {
                activeCell = _app.ActiveCell as Excel.Range;
                sheet = activeCell?.Worksheet ?? _app.ActiveSheet as Excel.Worksheet;
                if (sheet == null || activeCell == null)
                {
                    return;
                }

                if (ctrlPressed && TryMoveUsingEnd(activeCell, sheet, rowDelta, columnDelta, shiftPressed))
                {
                    return;
                }

                int maxRows = Convert.ToInt32(sheet.Rows.Count);
                int maxCols = Convert.ToInt32(sheet.Columns.Count);

                int destRow = Math.Max(1, Math.Min(maxRows, activeCell.Row + rowDelta));
                int destCol = ResolveVisibleColumn(sheet, activeCell.Column, columnDelta, maxCols);

                target = sheet.Cells[destRow, destCol] as Excel.Range;
                RangeHelpers.SafeSelect(target);
            }
            catch
            {
                // ignore
            }
            finally
            {
                ReleaseCom(target);
                ReleaseCom(activeCell);
                ReleaseCom(sheet);
            }
        }

        private bool TryMoveUsingEnd(Excel.Range activeCell, Excel.Worksheet sheet, int rowDelta, int columnDelta, bool extendSelection)
        {
            Excel.XlDirection? direction = null;
            if (rowDelta != 0)
            {
                direction = rowDelta > 0 ? Excel.XlDirection.xlDown : Excel.XlDirection.xlUp;
            }
            else if (columnDelta != 0)
            {
                direction = columnDelta > 0 ? Excel.XlDirection.xlToRight : Excel.XlDirection.xlToLeft;
            }

            if (direction == null)
            {
                return false;
            }

            Excel.Range destination = null;
            Excel.Range unionRange = null;
            try
            {
                destination = activeCell.get_End(direction.Value);
                if (!RangeHelpers.IsRangeValid(destination))
                {
                    return false;
                }

                if (extendSelection)
                {
                    unionRange = sheet.Range[destination, activeCell];
                    RangeHelpers.SafeSelect(unionRange);
                }
                else
                {
                    RangeHelpers.SafeSelect(destination);
                }
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(unionRange);
                ReleaseCom(destination);
            }
        }

        public void ActivateAdjacentSheet(int steps, bool forward)
        {
            steps = Math.Max(1, steps);

            Excel.Workbook workbook = null;
            Excel.Worksheet current = null;
            Excel.Sheets sheets = null;

            try
            {
                workbook = _app.ActiveWorkbook;
                current = _app.ActiveSheet as Excel.Worksheet;
                if (workbook == null || current == null)
                {
                    return;
                }

                sheets = workbook.Worksheets;
                int total = sheets.Count;
                if (total == 0)
                {
                    return;
                }

                int index = current.Index;
                int guard = 0;

                while (steps > 0 && guard < total * 2)
                {
                    index = forward ? (index % total) + 1 : ((index - 2 + total) % total) + 1;
                    guard++;

                    var candidate = sheets[index] as Excel.Worksheet;
                    if (candidate == null)
                    {
                        continue;
                    }

                    bool visible = candidate.Visible == Excel.XlSheetVisibility.xlSheetVisible;
                    if (visible)
                    {
                        steps--;
                        if (steps == 0)
                        {
                            RangeHelpers.SafeActivateSheet(candidate);
                        }
                    }

                    ReleaseCom(candidate);
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                ReleaseCom(sheets);
                ReleaseCom(current);
                ReleaseCom(workbook);
            }
        }

        public void MoveActiveSheet(int steps, bool forward)
        {
            steps = Math.Max(1, steps);

            Excel.Workbook workbook = null;
            Excel.Worksheet current = null;
            Excel.Sheets sheets = null;

            try
            {
                workbook = _app.ActiveWorkbook;
                current = _app.ActiveSheet as Excel.Worksheet;
                if (workbook == null || current == null)
                {
                    return;
                }

                sheets = workbook.Worksheets;
                int total = sheets.Count;
                if (total < 2)
                {
                    return;
                }

                int index = current.Index;
                int guard = 0;
                int targetIndex = index;

                while (steps > 0 && guard < total * 2)
                {
                    targetIndex = forward ? (targetIndex % total) + 1 : ((targetIndex - 2 + total) % total) + 1;
                    guard++;

                    var candidate = sheets[targetIndex] as Excel.Worksheet;
                    if (candidate == null || candidate == current)
                    {
                        ReleaseCom(candidate);
                        continue;
                    }

                    bool visible = candidate.Visible == Excel.XlSheetVisibility.xlSheetVisible;
                    if (visible)
                    {
                        steps--;
                        if (steps == 0)
                        {
                            object before = Type.Missing;
                            object after = Type.Missing;
                            if (forward)
                            {
                                after = candidate;
                            }
                            else
                            {
                                before = candidate;
                            }

                            current.Move(before, after);
                        }
                    }

                    ReleaseCom(candidate);
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                ReleaseCom(sheets);
                ReleaseCom(current);
                ReleaseCom(workbook);
            }
        }

        public void DeleteActiveCellComment()
        {
            if (!TryGetActiveCell(out var cell))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                DeleteCommentAt(cell);
            }

            ReleaseCom(cell);
        }

        public void DeleteAllComments()
        {
            if (!TryGetActiveWorksheet(out var sheet))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                DeleteLegacyComments(sheet);
                DeleteThreadedComments(sheet);
            }
        }

        public void ToggleActiveCommentVisibility()
        {
            if (!TryGetActiveCell(out var cell))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                var legacy = cell.Comment;
                if (legacy != null)
                {
                    try { legacy.Visible = !legacy.Visible; } catch { }
                    ReleaseCom(legacy);
                    ReleaseCom(cell);
                    return;
                }

                try
                {
                    var threaded = GetThreadedComment(cell);
                    if (threaded != null)
                    {
                        ToggleThreadedCommentDisplay(threaded);
                        ReleaseCom(threaded);
                    }
                }
                catch
                {
                    // ignore threaded comment failures
                }
            }

            ReleaseCom(cell);
        }

        public void ShowActiveComment() => SetActiveCommentVisibility(true);

        public void HideActiveComment() => SetActiveCommentVisibility(false);

        private void SetActiveCommentVisibility(bool visible)
        {
            if (!TryGetActiveCell(out var cell))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                var legacy = cell.Comment;
                if (legacy != null)
                {
                    try { legacy.Visible = visible; } catch { }
                    ReleaseCom(legacy);
                }
            }

            ReleaseCom(cell);
        }

        public void ToggleAllCommentsVisibility()
        {
            using (new UiGuard(_app))
            {
                try { _app.CommandBars.ExecuteMso("ReviewShowAllComments"); }
                catch { }
            }
        }

        public void SetCommentIndicatorMode(int mode)
        {
            Excel.XlCommentDisplayMode target = Excel.XlCommentDisplayMode.xlCommentIndicatorOnly;
            switch (mode)
            {
                case 0:
                    target = Excel.XlCommentDisplayMode.xlNoIndicator;
                    break;
                case 1:
                    target = Excel.XlCommentDisplayMode.xlCommentIndicatorOnly;
                    break;
                case 2:
                    target = Excel.XlCommentDisplayMode.xlCommentAndIndicator;
                    break;
            }

            try { _app.DisplayCommentIndicator = target; }
            catch { }
        }

        public void NavigateComments(bool forward, int steps)
        {
            if (steps < 1)
            {
                steps = 1;
            }

            if (!TryGetActiveWorksheet(out var sheet))
            {
                return;
            }

            var anchors = CollectCommentAnchors(sheet);
            if (anchors.Count == 0)
            {
                return;
            }

            int currentRow = 1;
            int currentCol = 1;
            Excel.Range activeCell = null;
            try
            {
                activeCell = _app.ActiveCell as Excel.Range;
                if (activeCell != null)
                {
                    currentRow = activeCell.Row;
                    currentCol = activeCell.Column;
                }
            }
            catch
            {
                currentRow = 1;
                currentCol = 1;
            }
            finally
            {
                ReleaseCom(activeCell);
            }

            int index = FindAnchorIndex(anchors, currentRow, currentCol, forward);
            if (index < 0)
            {
                index = forward ? 0 : anchors.Count - 1;
            }

            for (int i = 1; i < steps; i++)
            {
                index = forward ? (index + 1) % anchors.Count : (index - 1 + anchors.Count) % anchors.Count;
            }

            using (new UiGuard(_app))
            {
                Excel.Range target = null;
                try
                {
                    target = sheet.Cells[anchors[index].Row, anchors[index].Column] as Excel.Range;
                    if (target != null)
                    {
                        RangeHelpers.SafeSelect(target);
                        ShowCommentIfExists(target);
                    }
                }
                finally
                {
                    ReleaseCom(target);
                }
            }
        }

        public void PasteEntireRows(Excel.Range source, int copies, bool pasteAfterActive)
        {
            if (!RangeHelpers.IsRangeValid(source))
            {
                return;
            }

            if (copies < 1)
            {
                copies = 1;
            }

            if (!TryGetActiveWorksheet(out var sheet))
            {
                return;
            }

            Excel.Range activeCell = null;
            try { activeCell = _app.ActiveCell as Excel.Range; }
            catch { activeCell = null; }

            if (activeCell == null)
            {
                return;
            }

            int rowBlock = source.Rows.Count;
            if (rowBlock < 1)
            {
                return;
            }

            int maxRow = Convert.ToInt32(sheet.Rows.Count);
            int startRow = activeCell.Row + (pasteAfterActive ? 0 : 0);
            startRow = Math.Max(1, Math.Min(startRow, maxRow));

            int available = maxRow - startRow + 1;
            int maxCopies = available / rowBlock;
            if (maxCopies < 1)
            {
                return;
            }

            if (copies > maxCopies)
            {
                copies = maxCopies;
            }

            using (new UiGuard(_app))
            {
                Excel.Range insertBlock = null;
                try
                {
                    int insertEnd = startRow + rowBlock * copies - 1;
                    insertBlock = sheet.Range[sheet.Rows[startRow], sheet.Rows[insertEnd]];
                    insertBlock.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                    for (int i = 0; i < copies; i++)
                    {
                        int destStart = startRow + i * rowBlock;
                        var dest = sheet.Range[sheet.Rows[destStart], sheet.Rows[destStart + rowBlock - 1]];
                        source.Copy(dest);
                        ReleaseCom(dest);
                    }

                    source.Copy();
                }
                finally
                {
                    ReleaseCom(insertBlock);
                    ReleaseCom(activeCell);
                }
            }
        }

        public void PasteEntireColumns(Excel.Range source, int copies, bool pasteAfterActive)
        {
            if (!RangeHelpers.IsRangeValid(source))
            {
                return;
            }

            if (copies < 1)
            {
                copies = 1;
            }

            if (!TryGetActiveWorksheet(out var sheet))
            {
                return;
            }

            Excel.Range activeCell = null;
            try { activeCell = _app.ActiveCell as Excel.Range; }
            catch { activeCell = null; }

            if (activeCell == null)
            {
                return;
            }

            int colBlock = source.Columns.Count;
            if (colBlock < 1)
            {
                return;
            }

            int maxCol = Convert.ToInt32(sheet.Columns.Count);
            int startCol = activeCell.Column + (pasteAfterActive ? 0 : 0);
            startCol = Math.Max(1, Math.Min(startCol, maxCol));

            int available = maxCol - startCol + 1;
            int maxCopies = available / colBlock;
            if (maxCopies < 1)
            {
                return;
            }

            if (copies > maxCopies)
            {
                copies = maxCopies;
            }

            using (new UiGuard(_app))
            {
                Excel.Range insertBlock = null;
                try
                {
                    int insertEnd = startCol + colBlock * copies - 1;
                    insertBlock = sheet.Range[sheet.Columns[startCol], sheet.Columns[insertEnd]];
                    insertBlock.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                    for (int i = 0; i < copies; i++)
                    {
                        int destStart = startCol + i * colBlock;
                        var dest = sheet.Range[sheet.Columns[destStart], sheet.Columns[destStart + colBlock - 1]];
                        source.Copy(dest);
                        ReleaseCom(dest);
                    }

                    source.Copy();
                }
                finally
                {
                    ReleaseCom(insertBlock);
                    ReleaseCom(activeCell);
                }
            }
        }

        public void ResizeSelection(int up, int down, int left, int right)
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var selection))
            {
                return;
            }

            if (up == 0 && down == 0 && left == 0 && right == 0)
            {
                return;
            }

            Excel.Range baseRange = selection;
            Excel.Range activeCell = null;
            Excel.Range visible = null;
            Excel.Window window = null;
            int screenTop = 1, screenBottom = 1, screenLeft = 1, screenRight = 1;

            try { activeCell = _app.ActiveCell as Excel.Range; } catch { activeCell = null; }

            try
            {
                window = _app.ActiveWindow;
                visible = window?.VisibleRange;
                if (visible != null)
                {
                    screenTop = visible.Row;
                    screenLeft = visible.Column;
                    screenBottom = Math.Max(screenTop, screenTop + visible.Rows.Count - 2);
                    screenRight = Math.Max(screenLeft, screenLeft + visible.Columns.Count - 2);
                }
                else
                {
                    screenTop = selection.Row;
                    screenBottom = selection.Row + selection.Rows.Count - 1;
                    screenLeft = selection.Column;
                    screenRight = selection.Column + selection.Columns.Count - 1;
                }
            }
            catch
            {
                screenTop = selection.Row;
                screenBottom = selection.Row + selection.Rows.Count - 1;
                screenLeft = selection.Column;
                screenRight = selection.Column + selection.Columns.Count - 1;
            }

            int maxRow = Convert.ToInt32(selection.Worksheet.Rows.Count);
            int maxCol = Convert.ToInt32(selection.Worksheet.Columns.Count);

            int firstRow = selection.Row;
            int firstColumn = selection.Column;
            int lastRow = firstRow + selection.Rows.Count - 1;
            int lastColumn = firstColumn + selection.Columns.Count - 1;

            int rowCount = selection.Rows.Count;
            int colCount = selection.Columns.Count;

            if (up < 0 && -up >= rowCount)
            {
                down = -(rowCount + up) + 1;
                up = 0;
                baseRange = OffsetRange(baseRange, rowCount - 1, 0);
                baseRange = ResizeRange(baseRange, 1, baseRange.Columns.Count);
            }
            else if (down < 0 && -down >= rowCount)
            {
                up = -(rowCount + down) + 1;
                down = 0;
                baseRange = ResizeRange(baseRange, 1, baseRange.Columns.Count);
            }
            else if (left < 0 && -left >= colCount)
            {
                right = -(colCount + left) + 1;
                left = 0;
                baseRange = OffsetRange(baseRange, 0, colCount - 1);
                baseRange = ResizeRange(baseRange, baseRange.Rows.Count, 1);
            }
            else if (right < 0 && -right >= colCount)
            {
                left = -(colCount + right) + 1;
                right = 0;
                baseRange = ResizeRange(baseRange, baseRange.Rows.Count, 1);
            }

            if (up > 0 && firstRow <= up)
            {
                up = firstRow - 1;
            }
            else if (down > 0 && lastRow + down > maxRow)
            {
                down = maxRow - lastRow;
            }
            else if (left > 0 && firstColumn <= left)
            {
                left = firstColumn - 1;
            }
            else if (right > 0 && lastColumn + right > maxCol)
            {
                right = maxCol - lastColumn;
            }

            using (new UiGuard(_app))
            {
                if (up != 0)
                {
                    var resized = OffsetRange(baseRange, -up, 0);
                    var finalRange = ResizeRange(resized, baseRange.Rows.Count + up, baseRange.Columns.Count);
                    RangeHelpers.SafeSelect(finalRange);
                    SafeActivate(activeCell);
                    RestoreScroll(window, screenTop, screenLeft);
                    AdjustVerticalScroll(window, screenTop, screenBottom, firstRow - up);
                }
                else if (down != 0)
                {
                    var finalRange = ResizeRange(baseRange, baseRange.Rows.Count + down, baseRange.Columns.Count);
                    RangeHelpers.SafeSelect(finalRange);
                    SafeActivate(activeCell);
                    RestoreScroll(window, screenTop, screenLeft);
                    AdjustVerticalScroll(window, screenTop, screenBottom, lastRow + down);
                }
                else if (left != 0)
                {
                    var resized = OffsetRange(baseRange, 0, -left);
                    var finalRange = ResizeRange(resized, baseRange.Rows.Count, baseRange.Columns.Count + left);
                    RangeHelpers.SafeSelect(finalRange);
                    SafeActivate(activeCell);
                    RestoreScroll(window, screenTop, screenLeft);
                    AdjustHorizontalScroll(window, screenLeft, screenRight, firstColumn - left);
                }
                else if (right != 0)
                {
                    var finalRange = ResizeRange(baseRange, baseRange.Rows.Count, baseRange.Columns.Count + right);
                    RangeHelpers.SafeSelect(finalRange);
                    SafeActivate(activeCell);
                    RestoreScroll(window, screenTop, screenLeft);
                    AdjustHorizontalScroll(window, screenLeft, screenRight, lastColumn + right);
                }
            }

            ReleaseCom(visible);
            ReleaseCom(activeCell);
        }

        public void ScrollHalf(bool scrollUp, int repeatCount)
        {
            if (repeatCount < 1)
            {
                repeatCount = 1;
            }

            Excel.Window window = null;
            try { window = _app.ActiveWindow; }
            catch { window = null; }

            if (window == null)
            {
                return;
            }

            using (new UiGuard(_app))
            {
                object missing = Type.Missing;
                int visibleRows = GetVisibleRowCount(window);
                int half = Math.Max(1, visibleRows / 2);
                int largeScrolls = repeatCount / 2;

                for (int i = 0; i < largeScrolls; i++)
                {
                    if (scrollUp)
                    {
                        window.LargeScroll(missing, 1, missing, missing);
                    }
                    else
                    {
                        window.LargeScroll(1, missing, missing, missing);
                    }
                }

                if ((repeatCount & 1) == 1)
                {
                    if (scrollUp)
                    {
                        window.SmallScroll(missing, half, missing, missing);
                    }
                    else
                    {
                        window.SmallScroll(half, missing, missing, missing);
                    }
                }
            }

            EnsureActiveCellVisible();
        }

        public void ScrollHalfHorizontal(bool scrollLeft, int repeatCount)
        {
            if (repeatCount < 1)
            {
                repeatCount = 1;
            }

            Excel.Window window = null;
            try { window = _app.ActiveWindow; }
            catch { window = null; }

            if (window == null)
            {
                return;
            }

            using (new UiGuard(_app))
            {
                object missing = Type.Missing;
                int visibleColumns = GetVisibleColumnCount(window);
                int half = Math.Max(1, visibleColumns / 2);
                int largeScrolls = repeatCount / 2;

                for (int i = 0; i < largeScrolls; i++)
                {
                    if (scrollLeft)
                    {
                        window.LargeScroll(missing, missing, missing, 1);
                    }
                    else
                    {
                        window.LargeScroll(missing, missing, 1, missing);
                    }
                }

                if ((repeatCount & 1) == 1)
                {
                    if (scrollLeft)
                    {
                        window.SmallScroll(missing, missing, missing, half);
                    }
                    else
                    {
                        window.SmallScroll(missing, missing, half, missing);
                    }
                }
            }

            EnsureActiveCellVisible();
        }

        public void ScrollActiveRowToTop(double scrollOffsetPoints)
            => ScrollActiveRow(RowScrollPosition.Top, scrollOffsetPoints);

        public void ScrollActiveRowToBottom(double scrollOffsetPoints)
            => ScrollActiveRow(RowScrollPosition.Bottom, scrollOffsetPoints);

        public void ScrollActiveRowToMiddle()
            => ScrollActiveRow(RowScrollPosition.Middle, 0);

        public void ScrollActiveColumnToLeft()
            => ScrollActiveColumn(ColumnScrollPosition.Left);

        public void ScrollActiveColumnToRight()
            => ScrollActiveColumn(ColumnScrollPosition.Right);

        public void ScrollActiveColumnToCenter()
            => ScrollActiveColumn(ColumnScrollPosition.Center);

        public void EnsureActiveCellVisible()
        {
            Excel.Window window = null;
            Excel.Range visible = null;
            Excel.Range activeCell = null;
            try
            {
                window = _app.ActiveWindow;
                if (window == null)
                {
                    return;
                }

                visible = window.VisibleRange;
                activeCell = _app.ActiveCell as Excel.Range;
                if (visible == null || activeCell == null)
                {
                    return;
                }

                int visibleTop = visible.Row;
                int visibleLeft = visible.Column;
                int visibleBottom = visible.Row + visible.Rows.Count - 1;
                int visibleRight = visible.Column + visible.Columns.Count - 1;

                int targetRow = activeCell.Row;
                int targetCol = activeCell.Column;

                if (targetRow < visibleTop || targetRow > visibleBottom || targetCol < visibleLeft || targetCol > visibleRight)
                {
                    int clampRow = Math.Min(Math.Max(targetRow, visibleTop), visibleBottom);
                    int clampCol = Math.Min(Math.Max(targetCol, visibleLeft), visibleRight);
                    try
                    {
                        var cell = activeCell.Worksheet.Cells[clampRow, clampCol] as Excel.Range;
                        RangeHelpers.SafeSelect(cell);
                        cell?.Activate();
                        ReleaseCom(cell);
                    }
                    catch
                    {
                        // ignore
                    }

                    try { window.ScrollRow = visibleTop; } catch { }
                    try { window.ScrollColumn = visibleLeft; } catch { }
                }
            }
            finally
            {
                ReleaseCom(visible);
                ReleaseCom(activeCell);
            }
        }

        private void ScrollActiveRow(RowScrollPosition position, double scrollOffsetPoints)
        {
            if (!TryGetActiveWorksheet(out var sheet))
            {
                return;
            }

            Excel.Window window = null;
            Excel.Range activeCell = null;

            try
            {
                window = _app.ActiveWindow;
                activeCell = _app.ActiveCell as Excel.Range;
            }
            catch
            {
                window = null;
            }

            if (window == null || activeCell == null)
            {
                ReleaseCom(activeCell);
                return;
            }

            double usableHeight = GetRealUsableHeight(window);
            double cellTop = 0;
            double cellHeight = 0;
            try
            {
                cellTop = Convert.ToDouble(activeCell.Top);
                cellHeight = Convert.ToDouble(activeCell.Height);
            }
            catch
            {
                cellTop = 0;
                cellHeight = 0;
            }

            double point;
            double offset = Math.Max(0, scrollOffsetPoints);
            switch (position)
            {
                case RowScrollPosition.Top:
                    point = cellTop - GetDistanceAdjustedForZoom(offset, window);
                    break;
                case RowScrollPosition.Bottom:
                    point = cellTop + cellHeight - GetDistanceAdjustedForZoom(Math.Max(0, usableHeight - offset), window);
                    break;
                default:
                    point = cellTop + cellHeight / 2.0 - GetDistanceAdjustedForZoom(usableHeight, window) / 2.0;
                    break;
            }

            int scrollRow = PointToRow(sheet, window, point, position, offset);
            if (scrollRow > 0)
            {
                try { window.ScrollRow = scrollRow; } catch { }
            }

            ReleaseCom(activeCell);
        }

        private void ScrollActiveColumn(ColumnScrollPosition position)
        {
            if (!TryGetActiveWorksheet(out var sheet))
            {
                return;
            }

            Excel.Window window = null;
            Excel.Range activeCell = null;

            try
            {
                window = _app.ActiveWindow;
                activeCell = _app.ActiveCell as Excel.Range;
            }
            catch
            {
                window = null;
            }

            if (window == null || activeCell == null)
            {
                ReleaseCom(activeCell);
                return;
            }

            double usableWidth = GetRealUsableWidth(window);
            double cellLeft = 0;
            double cellWidth = 0;
            try
            {
                cellLeft = Convert.ToDouble(activeCell.Left);
                cellWidth = Convert.ToDouble(activeCell.Width);
            }
            catch
            {
                cellLeft = 0;
                cellWidth = 0;
            }

            double point;
            switch (position)
            {
                case ColumnScrollPosition.Left:
                    try { window.ScrollColumn = activeCell.Column; } catch { }
                    ReleaseCom(activeCell);
                    return;
                case ColumnScrollPosition.Right:
                    point = cellLeft + cellWidth - GetDistanceAdjustedForZoom(usableWidth, window);
                    break;
                default:
                    point = cellLeft + cellWidth / 2.0 - GetDistanceAdjustedForZoom(usableWidth, window) / 2.0;
                    break;
            }

            int column = PointToColumn(sheet, window, point, position);
            if (column > 0)
            {
                try { window.ScrollColumn = column; } catch { }
            }

            ReleaseCom(activeCell);
        }

        private double GetDistanceAdjustedForZoom(double value, Excel.Window window)
        {
            if (window == null)
            {
                return value;
            }

            double zoom;
            try { zoom = Convert.ToDouble(window.Zoom); }
            catch { zoom = 100d; }

            double rate;
            if (zoom > 90d && zoom < 110d)
            {
                rate = 1d;
            }
            else
            {
                rate = 103.32 / Math.Max(1d, zoom) - 0.05;
            }

            return value * rate;
        }

        private double GetRealUsableHeight(Excel.Window window)
        {
            if (window == null)
            {
                return 0;
            }

            double height;
            try { height = window.UsableHeight; }
            catch { height = 0; }

            try
            {
                if (window.DisplayHeadings)
                {
                    var sheet = _app.ActiveSheet as Excel.Worksheet;
                    height -= sheet?.StandardHeight ?? 0;
                }
            }
            catch
            {
                // ignore
            }

            return Math.Max(0, height);
        }

        private double GetRealUsableWidth(Excel.Window window)
        {
            if (window == null)
            {
                return 0;
            }

            double width;
            try { width = window.UsableWidth; }
            catch { width = 0; }

            bool headings;
            try { headings = window.DisplayHeadings; }
            catch { headings = false; }

            if (headings)
            {
                Excel.Range visible = null;
                Excel.Range lastCell = null;
                try
                {
                    visible = window.VisibleRange;
                    if (visible != null)
                    {
                        lastCell = visible.Cells[visible.Count] as Excel.Range;
                    }
                    int maxRow = lastCell?.Row ?? 0;
                    double headingWidth = 25d;
                    if (maxRow >= 1000)
                    {
                        headingWidth += 6.75 * (Convert.ToString(maxRow).Length - 3);
                    }

                    width -= headingWidth;
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseCom(lastCell);
                    ReleaseCom(visible);
                }
            }

            return Math.Max(0, width);
        }

        private int PointToRow(Excel.Worksheet sheet, Excel.Window window, double point, RowScrollPosition position, double scrollOffset)
        {
            int rowCount = GetRowCount(sheet);
            if (rowCount <= 0)
            {
                return 1;
            }

            double lastTop = GetRowTop(sheet, rowCount);
            if (point > lastTop)
            {
                return rowCount;
            }

            if (point <= 0)
            {
                return 1;
            }

            double avg = GetAverageRowHeight(window);
            int pred = (int)(point / avg) + 1;
            pred = Math.Max(1, Math.Min(rowCount, pred));
            double predTop = GetRowTop(sheet, pred);
            double diff = point - predTop;

            int l = pred;
            int h = pred;
            int i = 0;

            while (Math.Abs(diff) > double.Epsilon && i < 20)
            {
                int tmp = (int)Math.Round(diff / avg + 0.5) * (1 << i);
                if (tmp == 0)
                {
                    tmp = Math.Sign(diff) * (1 << i);
                }

                int candidate = pred + tmp;
                candidate = Math.Max(1, Math.Min(rowCount, candidate));

                if (diff < 0)
                {
                    h = l;
                    l = candidate;
                }
                else
                {
                    l = h;
                    h = candidate;
                }

                double lTop = GetRowTop(sheet, l);
                double hTop = GetRowTop(sheet, h);
                if (lTop <= point && point < hTop)
                {
                    break;
                }

                i++;
            }

            while (h - l >= 2)
            {
                int m = (int)Math.Round(l + (h - l) / 2.0 - 0.25, MidpointRounding.AwayFromZero);
                double mTop = GetRowTop(sheet, m);
                if (point < mTop)
                {
                    h = m;
                }
                else
                {
                    l = m;
                }
            }

            int result = l;
            double rowTop = GetRowTop(sheet, result);
            double rowHeight = GetRowHeight(sheet, result);

            switch (position)
            {
                case RowScrollPosition.Middle:
                    if ((point - rowTop) >= rowHeight / 2.0)
                    {
                        result++;
                    }
                    break;
                case RowScrollPosition.Top:
                    if (point > rowTop)
                    {
                        result++;
                    }
                    break;
                case RowScrollPosition.Bottom:
                    if (point - scrollOffset > rowTop)
                    {
                        result++;
                    }
                    break;
            }

            return Math.Max(1, Math.Min(rowCount, result));
        }

        private int PointToColumn(Excel.Worksheet sheet, Excel.Window window, double point, ColumnScrollPosition position)
        {
            int columnCount = GetColumnCount(sheet);
            if (point > GetColumnLeft(sheet, columnCount))
            {
                return columnCount;
            }

            if (point <= 0)
            {
                return 1;
            }

            double avg = GetAverageColumnWidth(window);
            int pred = (int)(point / avg) + 1;
            pred = Math.Max(1, Math.Min(columnCount, pred));
            double predLeft = GetColumnLeft(sheet, pred);
            double diff = point - predLeft;

            int l = pred;
            int h = pred;
            int i = 0;

            while (Math.Abs(diff) > double.Epsilon && i < 20)
            {
                int tmp = (int)Math.Round(diff / avg + 0.5) * (1 << i);
                if (tmp == 0)
                {
                    tmp = Math.Sign(diff) * (1 << i);
                }

                int candidate = pred + tmp;
                candidate = Math.Max(1, Math.Min(columnCount, candidate));

                if (diff < 0)
                {
                    h = l;
                    l = candidate;
                }
                else
                {
                    l = h;
                    h = candidate;
                }

                double lLeft = GetColumnLeft(sheet, l);
                double hLeft = GetColumnLeft(sheet, h);
                if (lLeft <= point && point < hLeft)
                {
                    break;
                }

                i++;
            }

            while (h - l >= 2)
            {
                int m = (int)Math.Round(l + (h - l) / 2.0 - 0.25, MidpointRounding.AwayFromZero);
                double mLeft = GetColumnLeft(sheet, m);
                if (point < mLeft)
                {
                    h = m;
                }
                else
                {
                    l = m;
                }
            }

            int result = l;
            double colLeft = GetColumnLeft(sheet, result);
            double colWidth = GetColumnWidth(sheet, result);

            switch (position)
            {
                case ColumnScrollPosition.Center:
                    if ((point - colLeft) >= colWidth / 2.0)
                    {
                        result++;
                    }
                    break;
                case ColumnScrollPosition.Right:
                    if (point > colLeft)
                    {
                        result++;
                    }
                    break;
            }

            return Math.Max(1, Math.Min(columnCount, result));
        }

        private int GetRowCount(Excel.Worksheet sheet)
        {
            try { return Convert.ToInt32(sheet.Rows.Count); }
            catch { return 1048576; }
        }

        private int GetColumnCount(Excel.Worksheet sheet)
        {
            try { return Convert.ToInt32(sheet.Columns.Count); }
            catch { return 16384; }
        }

        private double GetRowTop(Excel.Worksheet sheet, int rowIndex)
        {
            Excel.Range row = null;
            try
            {
                row = sheet.Rows[rowIndex] as Excel.Range;
                return row == null ? 0 : Convert.ToDouble(row.Top);
            }
            catch
            {
                return 0;
            }
            finally
            {
                ReleaseCom(row);
            }
        }

        private double GetRowHeight(Excel.Worksheet sheet, int rowIndex)
        {
            Excel.Range row = null;
            try
            {
                row = sheet.Rows[rowIndex] as Excel.Range;
                return row == null ? 0 : Convert.ToDouble(row.Height);
            }
            catch
            {
                return 0;
            }
            finally
            {
                ReleaseCom(row);
            }
        }

        private double GetColumnLeft(Excel.Worksheet sheet, int columnIndex)
        {
            Excel.Range column = null;
            try
            {
                column = sheet.Columns[columnIndex] as Excel.Range;
                return column == null ? 0 : Convert.ToDouble(column.Left);
            }
            catch
            {
                return 0;
            }
            finally
            {
                ReleaseCom(column);
            }
        }

        private double GetColumnWidth(Excel.Worksheet sheet, int columnIndex)
        {
            Excel.Range column = null;
            try
            {
                column = sheet.Columns[columnIndex] as Excel.Range;
                return column == null ? 0 : Convert.ToDouble(column.Width);
            }
            catch
            {
                return 0;
            }
            finally
            {
                ReleaseCom(column);
            }
        }

        private double GetAverageRowHeight(Excel.Window window)
        {
            Excel.Range visible = null;
            try
            {
                visible = window.VisibleRange;
                if (visible == null)
                {
                    return 15d;
                }

                double height = Convert.ToDouble(visible.Height);
                int rows = Math.Max(1, Convert.ToInt32(visible.Rows.Count));
                return height / rows;
            }
            catch
            {
                return 15d;
            }
            finally
            {
                ReleaseCom(visible);
            }
        }

        private double GetAverageColumnWidth(Excel.Window window)
        {
            Excel.Range visible = null;
            try
            {
                visible = window.VisibleRange;
                if (visible == null)
                {
                    return 8.43;
                }

                double width = Convert.ToDouble(visible.Width);
                int columns = Math.Max(1, Convert.ToInt32(visible.Columns.Count));
                return width / columns;
            }
            catch
            {
                return 8.43;
            }
            finally
            {
                ReleaseCom(visible);
            }
        }

        private int GetVisibleColumnCount(Excel.Window window)
        {
            Excel.Range visible = null;
            try
            {
                visible = window.VisibleRange;
                if (visible == null)
                {
                    return 1;
                }

                return Math.Max(1, visible.Columns.Count);
            }
            catch
            {
                return 1;
            }
            finally
            {
                ReleaseCom(visible);
            }
        }

        private bool TryGetActiveWorksheet(out Excel.Worksheet sheet)
        {
            sheet = null;
            try
            {
                sheet = _app.ActiveSheet as Excel.Worksheet;
            }
            catch
            {
                sheet = null;
            }

            return sheet != null;
        }

        private Excel.Range OffsetRange(Excel.Range range, int rowOffset, int columnOffset)
        {
            if (range == null)
            {
                return null;
            }

            try
            {
                return range.Offset[rowOffset, columnOffset];
            }
            catch
            {
                return range;
            }
        }

        private Excel.Range ResizeRange(Excel.Range range, int rows, int columns)
        {
            if (range == null)
            {
                return null;
            }

            try
            {
                return range.Resize[rows, columns];
            }
            catch
            {
                return range;
            }
        }

        private void SafeActivate(Excel.Range cell)
        {
            if (!RangeHelpers.IsRangeValid(cell))
            {
                return;
            }

            try
            {
                cell.Activate();
            }
            catch
            {
                // ignore
            }
        }

        private void RestoreScroll(Excel.Window window, int scrollRow, int scrollColumn)
        {
            if (window == null)
            {
                return;
            }

            try { window.ScrollRow = scrollRow; } catch { }
            try { window.ScrollColumn = scrollColumn; } catch { }
        }

        private void AdjustVerticalScroll(Excel.Window window, int screenTop, int screenBottom, int anchorRow)
        {
            if (window == null)
            {
                return;
            }

            object missing = Type.Missing;
            if (screenTop > anchorRow)
            {
                int delta = screenTop - anchorRow;
                if (delta > 0)
                {
                    try { window.SmallScroll(delta, missing, missing, missing); } catch { }
                }
            }
            else if (screenBottom < anchorRow)
            {
                int delta = anchorRow - screenBottom;
                if (delta > 0)
                {
                    try { window.SmallScroll(missing, delta, missing, missing); } catch { }
                }
            }
        }

        private void AdjustHorizontalScroll(Excel.Window window, int screenLeft, int screenRight, int anchorColumn)
        {
            if (window == null)
            {
                return;
            }

            object missing = Type.Missing;
            if (screenLeft > anchorColumn)
            {
                int delta = screenLeft - anchorColumn;
                if (delta > 0)
                {
                    try { window.SmallScroll(missing, missing, missing, delta); } catch { }
                }
            }
            else if (screenRight < anchorColumn)
            {
                int delta = anchorColumn - screenRight;
                if (delta > 0)
                {
                    try { window.SmallScroll(missing, missing, delta, missing); } catch { }
                }
            }
        }

        private int GetVisibleRowCount(Excel.Window window)
        {
            Excel.Range range = null;
            try
            {
                range = window.VisibleRange;
                if (range == null)
                {
                    return 1;
                }

                return Math.Max(1, range.Rows.Count);
            }
            catch
            {
                return 1;
            }
            finally
            {
                ReleaseCom(range);
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

        private bool TryGetActiveCell(out Excel.Range cell)
        {
            cell = null;
            try
            {
                cell = _app.ActiveCell as Excel.Range;
            }
            catch
            {
                cell = null;
            }

            return cell != null;
        }

        private sealed class CommentAnchor
        {
            public CommentAnchor(int row, int column)
            {
                Row = row;
                Column = column;
            }

            public int Row { get; }
            public int Column { get; }
        }

        private void DeleteCommentAt(Excel.Range cell)
        {
            if (cell == null)
            {
                return;
            }

            Excel.Comment legacy = null;
            try
            {
                legacy = cell.Comment;
                legacy?.Delete();
            }
            catch
            {
                // ignore
            }
            finally
            {
                ReleaseCom(legacy);
            }

            try
            {
                var threaded = GetThreadedComment(cell);
                if (threaded != null)
                {
                    InvokeDelete(threaded);
                    Marshal.FinalReleaseComObject(threaded);
                }
            }
            catch
            {
                // ignore
            }
        }

        private void DeleteLegacyComments(Excel.Worksheet sheet)
        {
            Excel.Comments comments = null;
            try
            {
                comments = sheet.Comments;
                if (comments == null)
                {
                    return;
                }

                foreach (Excel.Comment comment in comments)
                {
                    try { comment?.Delete(); }
                    catch { }
                    finally { ReleaseCom(comment); }
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                ReleaseCom(comments);
            }
        }

        private void DeleteThreadedComments(Excel.Worksheet sheet)
        {
            try
            {
                var collection = sheet.GetType().InvokeMember("CommentsThreaded", BindingFlags.GetProperty, null, sheet, null);
                if (collection == null)
                {
                    return;
                }

                var type = collection.GetType();
                int count = 0;
                try { count = (int)type.InvokeMember("Count", BindingFlags.GetProperty, null, collection, null); }
                catch { }

                for (int i = count; i >= 1; i--)
                {
                    object threaded = null;
                    try
                    {
                        threaded = type.InvokeMember("Item", BindingFlags.GetProperty, null, collection, new object[] { i });
                        InvokeDelete(threaded);
                    }
                    catch
                    {
                        // ignore
                    }
                    finally
                    {
                        if (threaded != null)
                        {
                            Marshal.FinalReleaseComObject(threaded);
                        }
                    }
                }

                Marshal.FinalReleaseComObject(collection);
            }
            catch
            {
                // property not available
            }
        }

        private object GetThreadedComment(Excel.Range cell)
        {
            if (cell == null)
            {
                return null;
            }

            try
            {
                return cell.GetType().InvokeMember("CommentThreaded", BindingFlags.GetProperty, null, cell, null);
            }
            catch
            {
                return null;
            }
        }

        private Excel.Range GetParentRange(object comment)
        {
            if (comment == null)
            {
                return null;
            }

            try
            {
                return comment.GetType().InvokeMember("Parent", BindingFlags.GetProperty, null, comment, null) as Excel.Range;
            }
            catch
            {
                return null;
            }
        }

        private void ToggleThreadedCommentDisplay(object threaded)
        {
            if (threaded == null)
            {
                return;
            }

            try
            {
                var type = threaded.GetType();
                object current = type.InvokeMember("ShowAlways", BindingFlags.GetProperty, null, threaded, null);
                bool state = current is bool b && b;
                type.InvokeMember("ShowAlways", BindingFlags.SetProperty, null, threaded, new object[] { !state });
            }
            catch
            {
                // ignore if property not present
            }
        }

        private void InvokeDelete(object target)
        {
            if (target == null)
            {
                return;
            }

            try
            {
                target.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, target, null);
            }
            catch
            {
                // ignore
            }
        }

        private List<CommentAnchor> CollectCommentAnchors(Excel.Worksheet sheet)
        {
            var anchors = new List<CommentAnchor>();
            if (sheet == null)
            {
                return anchors;
            }

            Excel.Comments comments = null;
            try
            {
                comments = sheet.Comments;
                if (comments != null)
                {
                    foreach (Excel.Comment legacy in comments)
                    {
                        Excel.Range parent = null;
                        try
                        {
                            parent = legacy?.Parent as Excel.Range;
                            if (RangeHelpers.IsRangeValid(parent))
                            {
                                anchors.Add(new CommentAnchor(parent.Row, parent.Column));
                            }
                        }
                        catch
                        {
                            // ignore
                        }
                        finally
                        {
                            ReleaseCom(parent);
                            ReleaseCom(legacy);
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                ReleaseCom(comments);
            }

            try
            {
                var threadedCollection = sheet.GetType().InvokeMember("CommentsThreaded", BindingFlags.GetProperty, null, sheet, null);
                if (threadedCollection != null)
                {
                    var type = threadedCollection.GetType();
                    int count = 0;
                    try { count = (int)type.InvokeMember("Count", BindingFlags.GetProperty, null, threadedCollection, null); }
                    catch { }

                    for (int i = 1; i <= count; i++)
                    {
                        object threaded = null;
                        Excel.Range parent = null;
                        try
                        {
                            threaded = type.InvokeMember("Item", BindingFlags.GetProperty, null, threadedCollection, new object[] { i });
                            parent = GetParentRange(threaded);
                            if (RangeHelpers.IsRangeValid(parent))
                            {
                                anchors.Add(new CommentAnchor(parent.Row, parent.Column));
                            }
                        }
                        catch
                        {
                            // ignore
                        }
                        finally
                        {
                            ReleaseCom(parent);
                            if (threaded != null)
                            {
                                Marshal.FinalReleaseComObject(threaded);
                            }
                        }
                    }

                    Marshal.FinalReleaseComObject(threadedCollection);
                }
            }
            catch
            {
                // threaded comments not supported
            }

            anchors.Sort((a, b) =>
            {
                int row = a.Row.CompareTo(b.Row);
                return row != 0 ? row : a.Column.CompareTo(b.Column);
            });

            return anchors;
        }

        private int FindAnchorIndex(IReadOnlyList<CommentAnchor> anchors, int row, int column, bool forward)
        {
            if (anchors == null || anchors.Count == 0)
            {
                return -1;
            }

            for (int i = 0; i < anchors.Count; i++)
            {
                var anchor = anchors[i];
                if (forward)
                {
                    if (anchor.Row > row || (anchor.Row == row && anchor.Column > column))
                    {
                        return i;
                    }
                }
                else
                {
                    if (anchor.Row < row || (anchor.Row == row && anchor.Column < column))
                    {
                        return i;
                    }
                }
            }

            return -1;
        }

        private void ShowCommentIfExists(Excel.Range cell)
        {
            if (cell == null)
            {
                return;
            }

            Excel.Comment legacy = null;
            try
            {
                legacy = cell.Comment;
                if (legacy != null)
                {
                    legacy.Visible = true;
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                ReleaseCom(legacy);
            }
        }

        private bool TryGetRowTarget(out Excel.Range targetRows, out double currentHeight)
        {
            targetRows = null;
            currentHeight = 0;
            Excel.Range selection = null;
            Excel.Range firstRow = null;

            try
            {
                if (RangeHelpers.TryGetActiveRange(_app, out selection))
                {
                    firstRow = selection.Rows[1] as Excel.Range ?? selection.Cells[1, 1] as Excel.Range;
                    targetRows = selection.EntireRow;
                }
                else
                {
                    firstRow = _app.ActiveCell as Excel.Range;
                    targetRows = firstRow?.EntireRow;
                }

                if (firstRow == null || targetRows == null)
                {
                    return false;
                }

                currentHeight = Convert.ToDouble(firstRow.RowHeight);
                return true;
            }
            catch
            {
                ReleaseCom(targetRows);
                targetRows = null;
                return false;
            }
            finally
            {
                ReleaseCom(firstRow);
                ReleaseCom(selection);
            }
        }

        private bool TryGetColumnTarget(out Excel.Range targetColumns, out double currentWidth)
        {
            targetColumns = null;
            currentWidth = 0;
            Excel.Range selection = null;
            Excel.Range firstColumn = null;

            try
            {
                if (RangeHelpers.TryGetActiveRange(_app, out selection))
                {
                    firstColumn = selection.Columns[1] as Excel.Range ?? selection.Cells[1, 1] as Excel.Range;
                    targetColumns = selection.EntireColumn;
                }
                else
                {
                    firstColumn = _app.ActiveCell as Excel.Range;
                    targetColumns = firstColumn?.EntireColumn;
                }

                if (firstColumn == null || targetColumns == null)
                {
                    return false;
                }

                currentWidth = Convert.ToDouble(firstColumn.ColumnWidth);
                return true;
            }
            catch
            {
                ReleaseCom(targetColumns);
                targetColumns = null;
                return false;
            }
            finally
            {
                ReleaseCom(firstColumn);
                ReleaseCom(selection);
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

        private int ResolveVisibleColumn(Excel.Worksheet sheet, int startColumn, int columnDelta, int maxColumns)
        {
            if (sheet == null)
            {
                return Math.Max(1, Math.Min(maxColumns, startColumn + columnDelta));
            }

            if (columnDelta == 0)
            {
                int stationary = Math.Max(1, Math.Min(maxColumns, startColumn));
                if (IsColumnVisible(sheet, stationary))
                {
                    return stationary;
                }

                int fallback = FindNextVisibleColumn(sheet, stationary, 1, maxColumns);
                if (fallback == -1)
                {
                    fallback = FindNextVisibleColumn(sheet, stationary, -1, maxColumns);
                }

                return fallback == -1 ? stationary : fallback;
            }

            int direction = columnDelta > 0 ? 1 : -1;
            int remaining = Math.Abs(columnDelta);
            int column = startColumn;

            while (remaining > 0)
            {
                column += direction;
                if (column < 1 || column > maxColumns)
                {
                    column = Math.Max(1, Math.Min(maxColumns, column));
                    break;
                }

                if (IsColumnVisible(sheet, column))
                {
                    remaining--;
                }
            }

            if (!IsColumnVisible(sheet, column))
            {
                int fallback = FindNextVisibleColumn(sheet, column, direction, maxColumns);
                if (fallback == -1)
                {
                    fallback = FindNextVisibleColumn(sheet, column, -direction, maxColumns);
                }

                if (fallback != -1)
                {
                    column = fallback;
                }
                else
                {
                    column = startColumn;
                }
            }

            return Math.Max(1, Math.Min(maxColumns, column));
        }

        private int FindNextVisibleColumn(Excel.Worksheet sheet, int startColumn, int direction, int maxColumns)
        {
            if (sheet == null || direction == 0)
            {
                return -1;
            }

            int column = startColumn;
            while (true)
            {
                column += direction;
                if (column < 1 || column > maxColumns)
                {
                    return -1;
                }

                if (IsColumnVisible(sheet, column))
                {
                    return column;
                }
            }
        }

        private bool IsColumnVisible(Excel.Worksheet sheet, int columnIndex)
        {
            Excel.Range column = null;
            try
            {
                if (sheet == null)
                {
                    return false;
                }

                column = sheet.Columns[columnIndex] as Excel.Range;
                if (column == null)
                {
                    return false;
                }

                return !Convert.ToBoolean(column.Hidden);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(column);
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

        private enum RowScrollPosition
        {
            Top = -1,
            Middle = 0,
            Bottom = 1
        }

        private enum ColumnScrollPosition
        {
            Left = -1,
            Center = 0,
            Right = 1
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
