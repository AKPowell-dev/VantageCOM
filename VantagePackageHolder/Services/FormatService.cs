using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace VantagePackageHolder
{
    internal sealed class FormatService
    {
        private readonly Excel.Application _app;
        private readonly ClipboardService _clipboard;
        private readonly PowerPointExporter _ppt;
        private readonly Dictionary<string, NumberFormatCycleState> _numberFormatStates = new Dictionary<string, NumberFormatCycleState>(StringComparer.Ordinal);
        private readonly Dictionary<string, bool> _borderCycleStates = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private long _selectionStamp = 1;
        private long _cycleFmtSelectionStampSeen;
        private long _borderCycleStampSeen;

        public FormatService(Excel.Application app, ClipboardService clipboard, PowerPointExporter ppt)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _clipboard = clipboard ?? throw new ArgumentNullException(nameof(clipboard));
            _ppt = ppt ?? throw new ArgumentNullException(nameof(ppt));
        }

        private static readonly BorderDescriptor BorderAroundDescriptor = new BorderDescriptor(
            "Around",
            new[]
            {
                Excel.XlBordersIndex.xlEdgeLeft,
                Excel.XlBordersIndex.xlEdgeTop,
                Excel.XlBordersIndex.xlEdgeBottom,
                Excel.XlBordersIndex.xlEdgeRight
            },
            new[] { "Around", "Left", "Top", "Bottom", "Right" });

        private static readonly BorderDescriptor BorderLeftDescriptor = new BorderDescriptor("Left", new[] { Excel.XlBordersIndex.xlEdgeLeft }, new[] { "Left" });
        private static readonly BorderDescriptor BorderTopDescriptor = new BorderDescriptor("Top", new[] { Excel.XlBordersIndex.xlEdgeTop }, new[] { "Top" });
        private static readonly BorderDescriptor BorderBottomDescriptor = new BorderDescriptor("Bottom", new[] { Excel.XlBordersIndex.xlEdgeBottom }, new[] { "Bottom" });
        private static readonly BorderDescriptor BorderRightDescriptor = new BorderDescriptor("Right", new[] { Excel.XlBordersIndex.xlEdgeRight }, new[] { "Right" });
        private static readonly BorderDescriptor BorderInsideHorizontalDescriptor = new BorderDescriptor("InsideHorizontal", new[] { Excel.XlBordersIndex.xlInsideHorizontal }, new[] { "InsideHorizontal" });
        private static readonly BorderDescriptor BorderInsideVerticalDescriptor = new BorderDescriptor("InsideVertical", new[] { Excel.XlBordersIndex.xlInsideVertical }, new[] { "InsideVertical" });
        private static readonly BorderDescriptor BorderInsideBothDescriptor = new BorderDescriptor(
            "InsideBoth",
            new[]
            {
                Excel.XlBordersIndex.xlInsideHorizontal,
                Excel.XlBordersIndex.xlInsideVertical
            },
            new[] { "InsideBoth" });
        private static readonly BorderDescriptor BorderDiagonalUpDescriptor = new BorderDescriptor("DiagonalUp", new[] { Excel.XlBordersIndex.xlDiagonalUp }, new[] { "DiagonalUp" });
        private static readonly BorderDescriptor BorderDiagonalDownDescriptor = new BorderDescriptor("DiagonalDown", new[] { Excel.XlBordersIndex.xlDiagonalDown }, new[] { "DiagonalDown" });
        private static readonly BorderDescriptor BorderAllDescriptor = new BorderDescriptor(
            "All",
            new[]
            {
                Excel.XlBordersIndex.xlEdgeLeft,
                Excel.XlBordersIndex.xlEdgeTop,
                Excel.XlBordersIndex.xlEdgeBottom,
                Excel.XlBordersIndex.xlEdgeRight,
                Excel.XlBordersIndex.xlInsideHorizontal,
                Excel.XlBordersIndex.xlInsideVertical
            },
            new[] { "Around", "Left", "Top", "Bottom", "Right", "InsideHorizontal", "InsideVertical", "InsideBoth", "All" });

        private static readonly IReadOnlyDictionary<string, BorderDescriptor> BorderDescriptors =
            new Dictionary<string, BorderDescriptor>(StringComparer.OrdinalIgnoreCase)
            {
                ["Around"] = BorderAroundDescriptor,
                ["Left"] = BorderLeftDescriptor,
                ["Top"] = BorderTopDescriptor,
                ["Bottom"] = BorderBottomDescriptor,
                ["Right"] = BorderRightDescriptor,
                ["InsideHorizontal"] = BorderInsideHorizontalDescriptor,
                ["InsideVertical"] = BorderInsideVerticalDescriptor,
                ["InsideBoth"] = BorderInsideBothDescriptor,
                ["DiagonalUp"] = BorderDiagonalUpDescriptor,
                ["DiagonalDown"] = BorderDiagonalDownDescriptor,
                ["All"] = BorderAllDescriptor
            };

        public void PasteExact()
        {
            using (new UiGuard(_app, hideStatusBar: true))
            {
                var src = _clipboard.GetCopyRange();
                if (!RangeHelpers.IsRangeValid(src))
                {
                    ExecuteMso("Paste");
                    return;
                }

                if (!(GetActiveRange() is Excel.Range dest))
                {
                    return;
                }

                int rowsCount = src.Rows.Count;
                int colsCount = src.Columns.Count;

                if (dest.Count == 1)
                {
                    dest = dest.Resize[rowsCount, colsCount];
                }
                else if (dest.Rows.Count != rowsCount || dest.Columns.Count != colsCount)
                {
                    dest = dest.Resize[rowsCount, colsCount];
                }

                try
                {
                    dest.Formula = src.Formula;
                    dest.NumberFormat = src.NumberFormat;
                    src.Copy();
                    dest.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                    if (_app.CutCopyMode == Excel.XlCutCopyMode.xlCopy)
                    {
                        src.Copy();
                    }

                    RunMacroIfExists("ClipboardRefresh");
                }
                catch
                {
                    ExecuteMso("Paste");
                }
            }
        }

        public void PasteCondensed()
        {
            using (new UiGuard(_app, hideStatusBar: true))
            {
                var src = _clipboard.GetCopyRange();
                if (!RangeHelpers.IsRangeValid(src))
                {
                    ExecuteMso("Paste");
                    return;
                }

                if (!(GetActiveRange() is Excel.Range destAnchor))
                {
                    return;
                }

                var rowKeep = CollectNonEmptyIndexes(src, isRow: true);
                var colKeep = CollectNonEmptyIndexes(src, isRow: false);

                if (rowKeep.Count == 0 || colKeep.Count == 0)
                {
                    return;
                }

                Excel.Range dest = destAnchor.Resize[rowKeep.Count, colKeep.Count];

                for (int r = 0; r < rowKeep.Count; r++)
                {
                    for (int c = 0; c < colKeep.Count; c++)
                    {
                        var srcCell = src.Cells[rowKeep[r], colKeep[c]] as Excel.Range;
                        var destCell = dest.Cells[r + 1, c + 1] as Excel.Range;
                        if (srcCell == null || destCell == null)
                        {
                            continue;
                        }

                        try
                        {
                            destCell.Formula = srcCell.Formula;
                            CopyCellPresentation(srcCell, destCell);
                        }
                        finally
                        {
                            ReleaseIfNeeded(srcCell);
                            ReleaseIfNeeded(destCell);
                        }
                    }
                }

                if (_app.CutCopyMode == Excel.XlCutCopyMode.xlCopy)
                {
                    src.Copy();
                }

                RunMacroIfExists("ClipboardRefresh");
            }
        }

        public void SmartFillRight()
        {
            using (new UiGuard(_app, hideStatusBar: true))
            {
                if (!(GetActiveRange() is Excel.Range selection))
                {
                    return;
                }

                var ws = selection.Worksheet;
                int startRow = selection.Row;
                int startCol = selection.Column;
                Excel.Range unionRange = null;

                for (int r = 1; r <= selection.Rows.Count; r++)
                {
                    Excel.Range sourceCell = null;
                    Excel.Range rowSpan = null;
                    Excel.Range targetRange = null;

                    try
                    {
                        sourceCell = selection.Cells[r, 1] as Excel.Range;
                        if (sourceCell == null)
                        {
                            continue;
                        }

                        int rowIndex = startRow + r - 1;
                        int lastCol = ComputeLastColFromNearestRow(ws, rowIndex, startCol + 1, 50, rowIndex, rowIndex);
                        if (lastCol <= startCol)
                        {
                            continue;
                        }

                        rowSpan = ws.Range[ws.Cells[rowIndex, startCol], ws.Cells[rowIndex, lastCol]];
                        targetRange = ws.Range[ws.Cells[rowIndex, startCol + 1], ws.Cells[rowIndex, lastCol]];

                                                // Borderless: assign formulas/values and format only
                        bool hasFormula = false;
                        try { hasFormula = (bool)sourceCell.HasFormula; } catch { }
                        try {
                            if (hasFormula) {
                                targetRange.Formula = sourceCell.Formula;
                            } else {
                                targetRange.Value2 = sourceCell.Value2;
                            }
                        } catch { }
                        ApplySourceFormatting(sourceCell, targetRange);
                        ApplyTopBottomBorders(sourceCell, targetRange);

                        if (unionRange == null)
                        {
                            unionRange = rowSpan;
                            rowSpan = null; // ownership transferred
                        }
                        else
                        {
                            var merged = _app.Union(unionRange, rowSpan);
                            ReleaseIfNeeded(unionRange);
                            unionRange = merged;
                            ReleaseIfNeeded(rowSpan);
                            rowSpan = null;
                        }
                    }
                    finally
                    {
                        ReleaseIfNeeded(sourceCell);
                        ReleaseIfNeeded(rowSpan);
                        ReleaseIfNeeded(targetRange);
                    }
                }

                RangeHelpers.SafeActivateSheet(ws);
                SelectRangeSafe(unionRange ?? selection);
                ReleaseIfNeeded(unionRange);
            }
        }

        public void SmartFormatRight()
        {
            using (new UiGuard(_app, hideStatusBar: true))
            {
                if (!(GetActiveRange() is Excel.Range selection))
                {
                    return;
                }

                var ws = selection.Worksheet;
                int startRow = selection.Row;
                int startCol = selection.Column;
                Excel.Range formattedUnion = null;

                try
                {
                    for (int r = 1; r <= selection.Rows.Count; r++)
                    {
                        Excel.Range sourceCell = null;
                        Excel.Range rowSpan = null;
                        Excel.Range targetRange = null;

                        try
                        {
                            sourceCell = selection.Cells[r, 1] as Excel.Range;
                            if (sourceCell == null)
                            {
                                continue;
                            }

                            int rowIndex = startRow + r - 1;
                            int lastCol = Math.Max(startCol, ComputeLastColFromNearestRow(ws, rowIndex, startCol + 1, 50, rowIndex, rowIndex));

                            rowSpan = ws.Range[ws.Cells[rowIndex, startCol], ws.Cells[rowIndex, lastCol]];
                            if (lastCol > startCol)
                            {
                                targetRange = ws.Range[ws.Cells[rowIndex, startCol + 1], ws.Cells[rowIndex, lastCol]];
                            }

                            ApplyFinanceRowFormatting(rowSpan, targetRange, sourceCell);
                            formattedUnion = MergeIntoUnion(formattedUnion, rowSpan);
                            rowSpan = null;
                        }
                        finally
                        {
                            ReleaseIfNeeded(sourceCell);
                            ReleaseIfNeeded(rowSpan);
                            ReleaseIfNeeded(targetRange);
                        }
                    }
                }
                finally
                {
                    RangeHelpers.SafeActivateSheet(ws);
                    SelectRangeSafe(formattedUnion ?? selection);
                    ReleaseIfNeeded(formattedUnion);
                }
            }
        }

        public void OutlineSelectionHighlight()
        {
            using (new UiGuard(_app, hideStatusBar: true))
            {
                if (!(GetActiveRange() is Excel.Range selection))
                {
                    return;
                }

                var ws = selection.Worksheet;
                int firstRow = selection.Row;
                int firstCol = selection.Column;
                int lastRow = firstRow + selection.Rows.Count - 1;
                int lastCol = firstCol + selection.Columns.Count - 1;
                bool skipTop = firstRow == 1;
                bool skipLeft = firstCol == 1;
                int highlightColor = ColorTranslator.ToOle(Color.FromArgb(0, 32, 96));
                var processed = new HashSet<string>(StringComparer.Ordinal);

                if (!skipTop)
                {
                    for (int col = firstCol; col <= lastCol; col++)
                    {
                        HighlightOutlineCell(ws, firstRow, col, highlightColor, firstRow, lastRow, firstCol, lastCol, skipTop, skipLeft, processed);
                    }
                }

                for (int col = firstCol; col <= lastCol; col++)
                {
                    HighlightOutlineCell(ws, lastRow, col, highlightColor, firstRow, lastRow, firstCol, lastCol, skipTop, skipLeft, processed);
                }

                if (!skipLeft)
                {
                    for (int row = firstRow; row <= lastRow; row++)
                    {
                        HighlightOutlineCell(ws, row, firstCol, highlightColor, firstRow, lastRow, firstCol, lastCol, skipTop, skipLeft, processed);
                    }
                }

                for (int row = firstRow; row <= lastRow; row++)
                {
                    HighlightOutlineCell(ws, row, lastCol, highlightColor, firstRow, lastRow, firstCol, lastCol, skipTop, skipLeft, processed);
                }

                Excel.Range leftColumn = null;
                Excel.Range rightColumn = null;
                try
                {
                    leftColumn = ws.Columns[firstCol] as Excel.Range;
                    rightColumn = ws.Columns[lastCol] as Excel.Range;
                    leftColumn.ColumnWidth = 2;
                    rightColumn.ColumnWidth = 2;
                }
                catch
                {
                    // ignore adjustments
                }
                finally
                {
                    ReleaseIfNeeded(leftColumn);
                    ReleaseIfNeeded(rightColumn);
                }
            }
        }

        public void SmartFillDown()
        {
            using (new UiGuard(_app, hideStatusBar: true))
            {
                if (!(GetActiveRange() is Excel.Range selection))
                {
                    return;
                }

                var ws = selection.Worksheet;
                int startRow = selection.Row;
                int startCol = selection.Column;
                int scanStart = startRow + 1;
                int maxRow = Convert.ToInt32(ws.Rows.Count);
                int maxCol = Convert.ToInt32(ws.Columns.Count);
                Excel.Range filledUnion = null;
                Excel.Range originalCell = null;

                try
                {
                    originalCell = _app.ActiveCell as Excel.Range;
                }
                catch
                {
                    originalCell = null;
                }

                for (int c = 1; c <= selection.Columns.Count; c++)
                {
                    Excel.Range sourceCell = null;
                    Excel.Range columnSpan = null;
                    Excel.Range targetRange = null;
                    Excel.Range changedTargets = null;

                    try
                    {
                        int currentCol = startCol + c - 1;
                        if (currentCol < 1 || currentCol > maxCol)
                        {
                            continue;
                        }

                        sourceCell = selection.Cells[1, c] as Excel.Range;
                        if (sourceCell == null)
                        {
                            continue;
                        }

                        bool sourceHasFill = Convert.ToInt32(sourceCell.Interior.ColorIndex) != (int)Excel.XlColorIndex.xlColorIndexNone;
                        int sourceFillColor = sourceHasFill ? Convert.ToInt32(sourceCell.Interior.Color) : 0;

                        int leftDist = int.MaxValue;
                        int rightDist = int.MaxValue;
                        int leftCol = 0;
                        int rightCol = 0;

                        for (int offset = 1; offset <= 5; offset++)
                        {
                            int candidate = currentCol - offset;
                            if (candidate >= 1 && ColumnHasData(ws, candidate, scanStart, maxRow))
                            {
                                leftDist = offset;
                                leftCol = candidate;
                                break;
                            }
                        }

                        for (int offset = 1; offset <= 5; offset++)
                        {
                            int candidate = currentCol + offset;
                            if (candidate <= maxCol && ColumnHasData(ws, candidate, scanStart, maxRow))
                            {
                                rightDist = offset;
                                rightCol = candidate;
                                break;
                            }
                        }

                        int nearestCol = 0;
                        if (leftDist < rightDist)
                        {
                            nearestCol = leftCol;
                        }
                        else if (rightDist < leftDist)
                        {
                            nearestCol = rightCol;
                        }
                        else if (leftDist == rightDist && leftCol > 0 && rightCol > 0)
                        {
                            int leftLast = GetLastDataRowOrStart(ws, leftCol, scanStart, maxRow, startRow);
                            int rightLast = GetLastDataRowOrStart(ws, rightCol, scanStart, maxRow, startRow);
                            nearestCol = leftLast < rightLast ? leftCol : rightCol;
                        }

                        int lastRow = startRow;
                        if (nearestCol > 0)
                        {
                            lastRow = GetLastDataRowOrStart(ws, nearestCol, scanStart, maxRow, startRow);
                        }

                        for (int row = startRow + 1; row <= lastRow; row++)
                        {
                            var cell = ws.Cells[row, currentCol] as Excel.Range;
                            bool hasValue = HasCellValue(cell);
                            ReleaseIfNeeded(cell);
                            if (hasValue)
                            {
                                lastRow = row - 1;
                                break;
                            }
                        }

                        if (lastRow <= startRow)
                        {
                            continue;
                        }

                        columnSpan = ws.Range[ws.Cells[startRow, currentCol], ws.Cells[lastRow, currentCol]];
                        targetRange = ws.Range[ws.Cells[startRow + 1, currentCol], ws.Cells[lastRow, currentCol]];

                        BorderState[] leftBorders = null;
                        BorderState[] rightBorders = null;
                        BorderState insideVertical = default;
                        BorderState insideHorizontal = default;
                        if (targetRange != null)
                        {
                            SnapshotBorders(targetRange, out leftBorders, out rightBorders, out _, out _, out insideVertical, out insideHorizontal);
                        }

                        bool sourceHasFormula = false;
                        object sourceFormula = null;
                        object sourceValue = null;
                        try { sourceHasFormula = Convert.ToBoolean(sourceCell.HasFormula); } catch { sourceHasFormula = false; }
                        if (sourceHasFormula)
                        {
                            try { sourceFormula = sourceCell.Formula; } catch { sourceHasFormula = false; }
                        }
                        if (!sourceHasFormula)
                        {
                            try { sourceValue = sourceCell.Value2; } catch { sourceValue = null; }
                        }

                        for (int row = startRow + 1; row <= lastRow; row++)
                        {
                            Excel.Range cell = null;
                            try
                            {
                                cell = ws.Cells[row, currentCol] as Excel.Range;
                                if (cell == null)
                                {
                                    continue;
                                }

                                if (HasCellValue(cell))
                                {
                                    continue;
                                }

                                try
                                {
                                    if (sourceHasFormula && sourceFormula != null)
                                    {
                                        cell.Formula = sourceFormula;
                                    }
                                    else
                                    {
                                        cell.Value2 = sourceValue;
                                    }
                                }
                                catch
                                {
                                    // ignore copy issues for individual cells
                                }

                                changedTargets = MergeIntoUnion(changedTargets, cell);
                                cell = null;
                            }
                            finally
                            {
                                ReleaseIfNeeded(cell);
                            }
                        }

                        if (changedTargets == null)
                        {
                            continue;
                        }

                        ApplyDownFillFormatting(changedTargets, sourceCell, sourceHasFill, sourceFillColor);
                        if (targetRange != null)
                        {
                            RestoreLeftRightAndInside(targetRange, leftBorders, rightBorders, insideVertical, insideHorizontal);
                        }
                        ApplyTopBottomBorders(sourceCell, changedTargets);

                        filledUnion = MergeIntoUnion(filledUnion, columnSpan);
                        columnSpan = null;
                    }
                    finally
                    {
                        ReleaseIfNeeded(sourceCell);
                        ReleaseIfNeeded(columnSpan);
                        ReleaseIfNeeded(targetRange);
                        ReleaseIfNeeded(changedTargets);
                    }
                }

                if (filledUnion != null)
                {
                    SelectRangeSafe(filledUnion);
                }
                else if (RangeHelpers.IsRangeValid(originalCell))
                {
                    SelectRangeSafe(originalCell);
                }
                else
                {
                    RangeHelpers.SafeSelect(selection);
                }

                ReleaseIfNeeded(filledUnion);
                ReleaseIfNeeded(originalCell);
            }
        }

        public void WrapFormulaWithCircCheck()
        {
            using (new UiGuard(_app, hideStatusBar: true))
            {
                if (!(GetActiveRange() is Excel.Range selection))
                {
                    return;
                }

                object formulasObj;
                try
                {
                    formulasObj = selection.Formula;
                }
                catch
                {
                    return;
                }

                switch (formulasObj)
                {
                    case string text:
                        {
                            var updated = WrapCircFormulaText(text);
                            if (!ReferenceEquals(updated, text) && updated != text)
                            {
                                selection.Formula = updated;
                            }

                            break;
                        }
                    case object single when single != null && !(single is Array):
                        {
                            var converted = Convert.ToString(single);
                            var replaced = WrapCircFormulaText(converted);
                            if (!ReferenceEquals(replaced, converted) && replaced != converted)
                            {
                                selection.Formula = replaced;
                            }

                            break;
                        }
                    case object[,] grid:
                        {
                            bool changed = false;
                            int rowLower = grid.GetLowerBound(0);
                            int rowUpper = grid.GetUpperBound(0);
                            int colLower = grid.GetLowerBound(1);
                            int colUpper = grid.GetUpperBound(1);

                            for (int r = rowLower; r <= rowUpper; r++)
                            {
                                for (int c = colLower; c <= colUpper; c++)
                                {
                                    if (grid[r, c] is string formulaText)
                                    {
                                        var newText = WrapCircFormulaText(formulaText);
                                        if (!ReferenceEquals(newText, formulaText) && newText != formulaText)
                                        {
                                            grid[r, c] = newText;
                                            changed = true;
                                        }
                                    }
                                }
                            }

                            if (changed)
                            {
                                selection.Formula = grid;
                            }

                            break;
                        }
                }
            }
        }

        public bool CopySelectionAsPicturePrintSafe()
        {
            var selection = _app.Selection;
            if (selection == null)
            {
                return false;
            }

            if (!CopySelectionAsPicture(selection))
            {
                if (!CopySelectionAsPicture(ResolveFallbackCopyTarget(selection)))
                {
                    return false;
                }
            }

            return _clipboard.ClipboardHasContent();
        }

        public void CopyPasteSelectionToPowerPoint()
        {
            if (!CopySelectionAsPicturePrintSafe())
            {
                System.Windows.Forms.MessageBox.Show(
                    "Unable to copy selection as picture (print view). Try selecting a different range or chart.",
                    "Copy to PowerPoint",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            _ppt.PasteClipboardIntoActiveSlide();
        }

        public void FormatChartFg()
        {
            var chart = ResolveTargetChart();
            if (chart == null)
            {
                return;
            }

            using (new UiGuard(_app))
            {
                const string fontName = "Garamond";
                const int fontSize = 11;
                int black = ColorTranslator.ToOle(Color.Black);

                FormatChartAreaFont(chart, fontName, fontSize, black);
                FormatCategoryAxis(chart, fontName, fontSize, black);
                SuppressChartTitle(chart);
                SuppressChartBorder(chart);
                DisableAxisGridlines(chart, Excel.XlAxisType.xlCategory);
                DisableAxisGridlines(chart, Excel.XlAxisType.xlValue);
                FormatBarOutlinesIfNeeded(chart, black);
                FormatAxisTitleFont(chart, Excel.XlAxisType.xlCategory, fontName, fontSize, black);
                FormatAxisTitleFont(chart, Excel.XlAxisType.xlValue, fontName, fontSize, black);
                FormatLegendFont(chart, fontName, fontSize, black);
            }
        }
        // ===== Cycle state reset for formatting cycles =====
        private string _cycleFmtLastKey = string.Empty;
        private int _cycleFmtNextStyle = 1;

        public void ResetCycleState()
        {
            unchecked
            {
                _selectionStamp++;
                if (_selectionStamp <= 0)
                {
                    _selectionStamp = 1;
                }
            }

            _cycleFmtLastKey = string.Empty;
            _cycleFmtNextStyle = 1;
            _cycleFmtSelectionStampSeen = 0;
            _borderCycleStates.Clear();
            _borderCycleStampSeen = 0;
            _numberFormatStates.Clear();
        }

        // ===== Robust LockCellReference (absolute references) =====
        public void LockCellReference()
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var sel))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                object formulasObj;
                try { formulasObj = sel.Formula; } catch { return; }

                object ConvertToAbsolute(object f)
                {
                    try
                    {
                        var s = f as string;
                        if (!string.IsNullOrEmpty(s) && s[0] == '=')
                        {
                            object abs = _app.ConvertFormula(
                                s,
                                Excel.XlReferenceStyle.xlA1,
                                Excel.XlReferenceStyle.xlA1,
                                Excel.XlReferenceType.xlAbsolute);
                            return Convert.ToString(abs);
                        }
                    }
                    catch { }
                    return f;
                }

                if (formulasObj is object[,] grid)
                {
                    int r0 = grid.GetLowerBound(0), r1 = grid.GetUpperBound(0);
                    int c0 = grid.GetLowerBound(1), c1 = grid.GetUpperBound(1);
                    for (int r = r0; r <= r1; r++)
                        for (int c = c0; c <= c1; c++)
                            grid[r, c] = ConvertToAbsolute(grid[r, c]);
                    try { sel.Formula = grid; } catch { }
                }
                else if (formulasObj is string fs)
                {
                    var converted = Convert.ToString(ConvertToAbsolute(fs));
                    if (!string.IsNullOrEmpty(converted))
                    {
                        try { sel.Formula = converted; } catch { }
                    }
                }
            }
        }

        public void ClearFormatting()
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var sel)) return;
            using (new UiGuard(_app))
            {
                try { sel.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }
                try { sel.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }
                try { sel.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }
                try { sel.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }
                try { sel.Interior.Pattern = Excel.XlPattern.xlPatternNone; } catch { }
                try { sel.NumberFormat = "#,##0_);(#,##0);--_)"; } catch { }
                try { sel.Font.Bold = false; } catch { }
                try { sel.Font.Italic = false; } catch { }
                try { sel.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone; } catch { }
                try { sel.Font.Color = ColorTranslator.ToOle(Color.Black); } catch { }
            }
        }

        public void CycleFormatting()
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var sel)) return;
            using (new UiGuard(_app))
            {
                string key = RangeHelpers.BuildRangeKey(sel);
                bool selectionMoved = _cycleFmtSelectionStampSeen != _selectionStamp;
                if (selectionMoved || !string.Equals(key, _cycleFmtLastKey, StringComparison.Ordinal))
                {
                    _cycleFmtNextStyle = 1;
                }
                _cycleFmtLastKey = key;
                _cycleFmtSelectionStampSeen = _selectionStamp;

                Excel.Range firstCell = null;
                try
                {
                    firstCell = sel.Cells[1, 1] as Excel.Range;
                    int BLUE = ColorTranslator.ToOle(Color.FromArgb(0, 32, 96));
                    int RED = ColorTranslator.ToOle(Color.FromArgb(153, 0, 0));
                    int LIGHTBLUE = ColorTranslator.ToOle(Color.FromArgb(226, 234, 250));

                    int next = _cycleFmtNextStyle;
                    int firstColor = Convert.ToInt32(firstCell.Font.Color);
                    bool firstHasFill = Convert.ToInt32(firstCell.Interior.Pattern) != (int)Excel.XlPattern.xlPatternNone;
                    int firstFill = firstHasFill ? Convert.ToInt32(firstCell.Interior.Color) : -1;
                    if (firstColor == RED) next = 2;
                    else if (firstHasFill && firstFill == BLUE) next = 3;
                    else if (firstHasFill && firstFill == LIGHTBLUE) next = 0;
                    else next = 1;

                    sel.Font.Name = "Garamond";
                    switch (next)
                    {
                        case 1:
                            sel.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                            sel.Font.Color = RED;
                            sel.Font.Bold = true;
                            sel.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
                            sel.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            sel.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            break;
                        case 2:
                            sel.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                            sel.Interior.Color = BLUE;
                            sel.Font.Color = ColorTranslator.ToOle(Color.White);
                            sel.Font.Bold = true;
                            sel.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone;
                            sel.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            sel.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            break;
                        case 3:
                            sel.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                            sel.Interior.Color = LIGHTBLUE;
                            sel.Font.Color = ColorTranslator.ToOle(Color.Black);
                            sel.Font.Bold = true;
                            sel.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone;
                            var t = sel.Borders[Excel.XlBordersIndex.xlEdgeTop];
                            t.LineStyle = Excel.XlLineStyle.xlContinuous; t.Weight = Excel.XlBorderWeight.xlThin;
                            var b = sel.Borders[Excel.XlBordersIndex.xlEdgeBottom];
                            b.LineStyle = Excel.XlLineStyle.xlContinuous; b.Weight = Excel.XlBorderWeight.xlThin;
                            break;
                        default:
                            sel.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                            sel.Font.Color = ColorTranslator.ToOle(Color.Black);
                            sel.Font.Bold = false;
                            sel.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone;
                            sel.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            sel.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            break;
                    }
                    _cycleFmtNextStyle = next + 1; if (_cycleFmtNextStyle > 3) _cycleFmtNextStyle = 0;
                }
                finally { ReleaseIfNeeded(firstCell); }
            }
        }

        public void ToggleBorder(string targetKey, int lineStyle, int weight)
        {
            if (!TryGetBorderDescriptor(targetKey, out var descriptor))
            {
                return;
            }

            ToggleBorder(descriptor, lineStyle, weight);
        }

        public void DeleteBorder(string targetKey)
        {
            if (!TryGetBorderDescriptor(targetKey, out var descriptor))
            {
                return;
            }

            DeleteBorder(descriptor);
        }

        public void SetBorderColor(string targetKey, bool isNull, bool isThemeColor, int themeColor, double tintAndShade, int rgb)
        {
            if (!TryGetBorderDescriptor(targetKey, out var descriptor))
            {
                return;
            }

            var spec = new BorderColorSpec(isNull, isThemeColor, themeColor, tintAndShade, rgb);
            ApplyBorderColor(descriptor, spec);
        }

        public void ApplyInteriorColor(bool isNull, bool isThemeColor, int themeColor, double tintAndShade, int rgb)
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var sel)) return;
            using (new UiGuard(_app))
            {
                try
                {
                    var spec = new ColorSpec(isNull, isThemeColor, themeColor, themeColor, tintAndShade, rgb);
                    ApplyInterior(sel.Interior, spec);
                }
                catch
                {
                    // ignore
                }
            }
        }

        public bool ApplyFontColor(bool isNull, bool isThemeColor, int themeColor, int objectThemeColor, double tintAndShade, int rgb)
        {
            var spec = new ColorSpec(isNull, isThemeColor, themeColor, objectThemeColor, tintAndShade, rgb);
            using (new UiGuard(_app))
            {
                return TryApplyFontColorToSelection(spec);
            }
        }

        public void ApplyShapeFillColor(bool isNull, bool isThemeColor, int themeColor, double tintAndShade, int rgb)
        {
            using (new UiGuard(_app))
            {
                var spec = new ColorSpec(isNull, isThemeColor, themeColor, themeColor, tintAndShade, rgb);
                TryApplyShapeFill(spec);
            }
        }

        public void ApplyShapeFontColor(bool isNull, bool isThemeColor, int themeColor, int objectThemeColor, double tintAndShade, int rgb)
        {
            using (new UiGuard(_app))
            {
                var spec = new ColorSpec(isNull, isThemeColor, themeColor, objectThemeColor, tintAndShade, rgb);
                TryApplyShapeFont(spec);
            }
        }

        public void ApplyShapeBorderColor(bool isNull, bool isThemeColor, int themeColor, double tintAndShade, int rgb)
        {
            using (new UiGuard(_app))
            {
                var spec = new ColorSpec(isNull, isThemeColor, themeColor, themeColor, tintAndShade, rgb);
                TryApplyShapeBorder(spec);
            }
        }

        public bool ApplySmartFontColor(bool isNull, bool isThemeColor, int themeColor, int objectThemeColor, double tintAndShade, int rgb)
        {
            var spec = new ColorSpec(isNull, isThemeColor, themeColor, objectThemeColor, tintAndShade, rgb);
            using (new UiGuard(_app))
            {
                if (TryApplyFontColorToSelection(spec))
                {
                    return true;
                }

                return TryApplyShapeFont(spec);
            }
        }

        public void ApplySmartFillColor(bool isNull, bool isThemeColor, int themeColor, double tintAndShade, int rgb)
        {
            using (new UiGuard(_app))
            {
                var spec = new ColorSpec(isNull, isThemeColor, themeColor, themeColor, tintAndShade, rgb);
                if (RangeHelpers.TryGetActiveRange(_app, out var sel))
                {
                    try
                    {
                        ApplyInterior(sel.Interior, spec);
                        return;
                    }
                    catch
                    {
                        // fall through to shapes
                    }
                }

                TryApplyShapeFill(spec);
            }
        }

        public void CycleNumberFormat()
        {
            var formats = new[]
            {
                "#,##0_);(#,##0);--_)",
                "$#,##0_);($#,##0);$--_)",
                @"#,##0.0%_);(#,##0.0%);--\%_)",
                "#,##0.0x_);(#,##0.0x);--x_)",
                @"#,##0""bps""_);(#,##0""bps"");""--bps """,
                @"""On"";"""";""Off""",
                @"[>=1]""Yes"";""No"";""No""",
                @"[=1]0"" Year"";0"" Years""",
                @"""Year ""0; ""Year ""-0; ""Year 0""; """""
            };
            ApplyNumberFormatCycle(nameof(CycleNumberFormat), formats);
        }

        public void BinaryCycle()
        {
            var formats = new[]
            {
                @"[>=1]""Yes"";""No"";""No""",
                @"""On"";"""";""Off"""
            };
            ApplyNumberFormatCycle(nameof(BinaryCycle), formats);
        }

        public void YearDisplayCycle()
        {
            var formats = new[]
            {
                "yyyy",
                "mmm-yy"
            };
            ApplyNumberFormatCycle(nameof(YearDisplayCycle), formats);
        }

        public void NumberNarrativeCycle()
        {
            var formats = new[]
            {
                "#,##0_);(#,##0);--_)",
                "#,##0.0x_);(#,##0.0x);--x_)",
                @"[=1]0"" Year"";0"" Years""",
                @"""Year ""0; ""Year ""-0; ""Year 0""; """""
            };
            ApplyNumberFormatCycle(nameof(NumberNarrativeCycle), formats);
        }

        public void PercentCycle()
        {
            var formats = new[]
            {
                @"#,##0.0%_);(#,##0.0%);--\%_)",
                @"#,##0""bps""_);(#,##0""bps"");""--bps """
            };
            ApplyNumberFormatCycle(nameof(PercentCycle), formats);
        }

        public void CurrencyCycle()
        {
            string pound = char.ConvertFromUtf32(0x00A3);
            string euro = char.ConvertFromUtf32(0x20AC);
            var formats = new[]
            {
                "$#,##0_);($#,##0);$--_)",
                pound + "#,##0_);(" + pound + "#,##0);" + pound + "--_)",
                euro + "#,##0_);(" + euro + "#,##0);" + euro + "--_)"
            };
            ApplyNumberFormatCycle(nameof(CurrencyCycle), formats);
        }

        private void ApplyNumberFormatCycle(string cycleName, string[] formats)
        {
            if (formats == null || formats.Length == 0)
            {
                return;
            }

            if (!RangeHelpers.TryGetActiveRange(_app, out var sel))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                var key = RangeHelpers.BuildRangeKey(sel);
                if (!_numberFormatStates.TryGetValue(cycleName, out var state))
                {
                    state = new NumberFormatCycleState();
                    _numberFormatStates[cycleName] = state;
                }

                if (!string.Equals(key, state.LastSelectionKey, StringComparison.Ordinal))
                {
                    state.NextIndex = 0;
                }

                state.LastSelectionKey = key;
                string format = formats[state.NextIndex];
                bool success = false;
                try
                {
                    sel.NumberFormat = format;
                    success = true;
                }
                catch
                {
                    // ignore and fall back
                }

                if (!success)
                {
                    try
                    {
                        sel.NumberFormatLocal = format;
                    }
                    catch
                    {
                        // ignore
                    }
                }

                state.NextIndex = (state.NextIndex + 1) % formats.Length;
            }
        }

        public void FlipSign()
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var sel)) return;
            using (new UiGuard(_app))
            {
                foreach (Excel.Range cell in sel.Cells)
                {
                    try
                    {
                        if (Convert.ToBoolean(cell.HasArray)) { continue; }
                        if (Convert.ToBoolean(cell.HasFormula))
                        {
                            var v = cell.Value2;
                            if (v != null && double.TryParse(Convert.ToString(v), out _))
                            {
                                string f = Convert.ToString(cell.Formula);
                                if (f.StartsWith("=-(") && f.EndsWith(")")) cell.Formula = "=" + f.Substring(3, f.Length - 4);
                                else if (f.StartsWith("=")) cell.Formula = "=-(" + f.Substring(1) + ")";
                            }
                        }
                        else
                        {
                            var v = cell.Value2;
                            if (v != null && double.TryParse(Convert.ToString(v), out double num)) cell.Value2 = -num;
                        }
                    }
                    catch { }
                    finally { ReleaseIfNeeded(cell); }
                }
            }
        }

        public void ReverseSelectionOrder()
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var sel)) return;
            if (sel.Areas.Count > 1) return;
            long total = Convert.ToInt64(sel.Cells.CountLarge); if (total < 2) return;
            using (new UiGuard(_app))
            {
                var values = new object[total + 1];
                var formulas = new string[total + 1];
                var hasFormula = new bool[total + 1];
                int i = 1;
                foreach (Excel.Range cell in sel.Cells)
                {
                    try
                    {
                        if (Convert.ToBoolean(cell.MergeCells)) return;
                        if (Convert.ToBoolean(cell.HasArray)) return;
                        hasFormula[i] = Convert.ToBoolean(cell.HasFormula);
                        if (hasFormula[i]) formulas[i] = Convert.ToString(cell.Formula); else values[i] = cell.Value2;
                        i++;
                    }
                    catch { return; }
                    finally { ReleaseIfNeeded(cell); }
                }
                i = 1;
                foreach (Excel.Range cell in sel.Cells)
                {
                    int src = (int)total - i + 1;
                    try { if (hasFormula[src]) cell.Formula = formulas[src]; else cell.Value2 = values[src]; i++; }
                    catch { return; }
                    finally { ReleaseIfNeeded(cell); }
                }
            }
        }

        public void TrimConditionalFormatting()
        {
            Excel.Worksheet sheet = null;
            try
            {
                sheet = _app.ActiveSheet as Excel.Worksheet;
            }
            catch
            {
                sheet = null;
            }

            if (sheet == null)
            {
                return;
            }

            using (new UiGuard(_app))
            {
                try
                {
                    var conditions = sheet.Cells.FormatConditions;
                    if (conditions != null && conditions.Count > 0)
                    {
                        conditions.Delete();
                    }
                }
                catch
                {
                    // ignore
                }
            }
        }

        public void HighlightSelectionYellow()
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var range))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                range.Interior.Color = (int)Excel.XlRgbColor.rgbYellow;
            }
        }

        #region PasteCondensed helpers
        private List<int> CollectNonEmptyIndexes(Excel.Range src, bool isRow)
        {
            var keep = new List<int>();
            var worksheetFunction = _app.WorksheetFunction;
            int count = isRow ? src.Rows.Count : src.Columns.Count;

            for (int i = 1; i <= count; i++)
            {
                Excel.Range slice = isRow ? (Excel.Range)src.Rows[i] : (Excel.Range)src.Columns[i];
                try
                {
                    double hasData = worksheetFunction.CountA(slice);
                    if (hasData > 0)
                    {
                        keep.Add(i);
                    }
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseIfNeeded(slice);
                }
            }

            return keep;
        }

        private static void CopyCellPresentation(Excel.Range srcCell, Excel.Range destCell)
        {
            if (srcCell == null || destCell == null)
            {
                return;
            }

            try
            {
                destCell.NumberFormat = srcCell.NumberFormat;

                destCell.Font.Name = srcCell.Font.Name;
                destCell.Font.Size = srcCell.Font.Size;
                destCell.Font.Bold = srcCell.Font.Bold;
                destCell.Font.Italic = srcCell.Font.Italic;
                destCell.Font.Underline = srcCell.Font.Underline;
                destCell.Font.Color = srcCell.Font.Color;
                destCell.Font.Strikethrough = srcCell.Font.Strikethrough;

                int sourcePattern = Convert.ToInt32(srcCell.Interior.Pattern);
                int sourceColorIndex = Convert.ToInt32(srcCell.Interior.ColorIndex);
                if (sourcePattern == (int)Excel.XlPattern.xlPatternNone ||
                    sourceColorIndex == (int)Excel.XlColorIndex.xlColorIndexNone)
                {
                    destCell.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                    destCell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                }
                else
                {
                    destCell.Interior.Pattern = srcCell.Interior.Pattern;
                    destCell.Interior.PatternColorIndex = srcCell.Interior.PatternColorIndex;
                    destCell.Interior.Color = srcCell.Interior.Color;
                    destCell.Interior.TintAndShade = srcCell.Interior.TintAndShade;
                    destCell.Interior.PatternTintAndShade = srcCell.Interior.PatternTintAndShade;
                }

                destCell.HorizontalAlignment = srcCell.HorizontalAlignment;
                destCell.VerticalAlignment = srcCell.VerticalAlignment;
                destCell.WrapText = srcCell.WrapText;
                destCell.Orientation = srcCell.Orientation;
                destCell.AddIndent = srcCell.AddIndent;
                destCell.IndentLevel = srcCell.IndentLevel;
                destCell.ShrinkToFit = srcCell.ShrinkToFit;
                destCell.ReadingOrder = srcCell.ReadingOrder;

                foreach (Excel.XlBordersIndex borderIndex in new[]
                         {
                             Excel.XlBordersIndex.xlEdgeLeft,
                             Excel.XlBordersIndex.xlEdgeRight,
                             Excel.XlBordersIndex.xlEdgeTop,
                             Excel.XlBordersIndex.xlEdgeBottom
                         })
                {
                    var destBorder = destCell.Borders[borderIndex];
                    var srcBorder = srcCell.Borders[borderIndex];
                    destBorder.LineStyle = srcBorder.LineStyle;
                    if (Convert.ToInt32(destBorder.LineStyle) != (int)Excel.XlLineStyle.xlLineStyleNone)
                    {
                        destBorder.Weight = srcBorder.Weight;
                        destBorder.Color = srcBorder.Color;
                    }
                }
            }
            catch
            {
                // ignored
            }
        }
        #endregion

        #region Chart formatting helpers
        private Excel.Chart ResolveTargetChart()
        {
            try
            {
                var activeChart = _app.ActiveChart;
                if (activeChart != null)
                {
                    return activeChart;
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var selection = _app.Selection;
                switch (selection)
                {
                    case Excel.ChartObject chartObject:
                        return chartObject.Chart;
                    case Excel.Shape shape when shape.HasChart == Office.MsoTriState.msoTrue:
                        return shape.Chart;
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private void FormatChartAreaFont(Excel.Chart chart, string fontName, int fontSize, int color)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                var area = chart.ChartArea;
                area.Font.Name = fontName;
                area.Font.Size = fontSize;
                area.Font.Color = color;
            }
            catch
            {
                // ignore
            }
        }

        private void FormatCategoryAxis(Excel.Chart chart, string fontName, int fontSize, int color)
        {
            Excel.Axis axis = null;
            Excel.TickLabels tickLabels = null;
            try
            {
                axis = GetAxis(chart, Excel.XlAxisType.xlCategory);
                if (axis == null)
                {
                    return;
                }

                try
                {
                    tickLabels = axis.TickLabels;
                    if (tickLabels != null)
                    {
                        var font = tickLabels.Font;
                        font.Name = fontName;
                        font.Size = fontSize;
                        font.Color = color;
                    }
                }
                catch
                {
                    // ignore
                }

                axis.MajorTickMark = Excel.XlTickMark.xlTickMarkOutside;
                try
                {
                    var line = axis.Format.Line;
                    line.Visible = Office.MsoTriState.msoTrue;
                    line.ForeColor.RGB = color;
                }
                catch
                {
                    // ignore
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                ReleaseIfNeeded(tickLabels);
                ReleaseIfNeeded(axis);
            }
        }

        private void SuppressChartTitle(Excel.Chart chart)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                if (chart.HasTitle)
                {
                    chart.HasTitle = false;
                }
            }
            catch
            {
                // ignore
            }
        }

        private void SuppressChartBorder(Excel.Chart chart)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                chart.ChartArea.Format.Line.Visible = Office.MsoTriState.msoFalse;
            }
            catch
            {
                // ignore
            }
        }

        private void DisableAxisGridlines(Excel.Chart chart, Excel.XlAxisType axisType)
        {
            Excel.Axis axis = null;
            try
            {
                axis = GetAxis(chart, axisType);
                if (axis != null)
                {
                    axis.HasMajorGridlines = false;
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                ReleaseIfNeeded(axis);
            }
        }

        private void FormatAxisTitleFont(Excel.Chart chart, Excel.XlAxisType axisType, string fontName, int fontSize, int color)
        {
            Excel.Axis axis = null;
            try
            {
                axis = GetAxis(chart, axisType);
                if (axis == null || !axis.HasTitle)
                {
                    return;
                }

                var title = axis.AxisTitle;
                title.Font.Name = fontName;
                title.Font.Size = fontSize;
                title.Font.Color = color;
            }
            catch
            {
                // ignore
            }
            finally
            {
                ReleaseIfNeeded(axis);
            }
        }

        private void FormatLegendFont(Excel.Chart chart, string fontName, int fontSize, int color)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                if (!chart.HasLegend)
                {
                    return;
                }

                var legend = chart.Legend;
                legend.Font.Name = fontName;
                legend.Font.Size = fontSize;
                legend.Font.Color = color;
            }
            catch
            {
                // ignore
            }
        }

        private void FormatBarOutlinesIfNeeded(Excel.Chart chart, int color)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                var type = chart.ChartType;
                if (!IsBarOrColumnChart(type))
                {
                    return;
                }

                Excel.SeriesCollection seriesCollection = null;
                try
                {
                    seriesCollection = chart.SeriesCollection() as Excel.SeriesCollection;
                    if (seriesCollection == null)
                    {
                        return;
                    }

                    int seriesCount = seriesCollection.Count;
                    for (int i = 1; i <= seriesCount; i++)
                    {
                        Excel.Series series = null;
                        try
                        {
                            series = seriesCollection.Item(i);
                            var line = series.Format.Line;
                            line.Visible = Office.MsoTriState.msoTrue;
                            line.ForeColor.RGB = color;
                            line.Weight = 0.75f;
                        }
                        catch
                        {
                            // ignore
                        }
                        finally
                        {
                            ReleaseIfNeeded(series);
                        }
                    }
                }
                finally
                {
                    ReleaseIfNeeded(seriesCollection);
                }

                Excel.ChartGroup group = null;
                try
                {
                    group = chart.ChartGroups(1) as Excel.ChartGroup;
                    if (group != null)
                    {
                        group.GapWidth = 15;
                    }
                }
                catch
                {
                    // ignore
                }
                finally
                {
                    ReleaseIfNeeded(group);
                }
            }
            catch
            {
                // ignore
            }
        }

        private static bool IsBarOrColumnChart(Excel.XlChartType type)
        {
            switch (type)
            {
                case Excel.XlChartType.xlBarClustered:
                case Excel.XlChartType.xlBarStacked:
                case Excel.XlChartType.xlBarStacked100:
                case Excel.XlChartType.xlColumnClustered:
                case Excel.XlChartType.xlColumnStacked:
                case Excel.XlChartType.xlColumnStacked100:
                    return true;
                default:
                    return false;
            }
        }

        private Excel.Axis GetAxis(Excel.Chart chart, Excel.XlAxisType axisType)
        {
            if (chart == null)
            {
                return null;
            }

            try
            {
                return chart.Axes(axisType, Excel.XlAxisGroup.xlPrimary) as Excel.Axis;
            }
            catch
            {
                return null;
            }
        }
        #endregion

        #region SmartFill helpers
        private static void SnapshotBorders(
            Excel.Range fullRange,
            out BorderState[] left,
            out BorderState[] right,
            out BorderState[] top,
            out BorderState[] bottom,
            out BorderState insideVertical,
            out BorderState insideHorizontal)
        {
            left = Array.Empty<BorderState>();
            right = Array.Empty<BorderState>();
            top = Array.Empty<BorderState>();
            bottom = Array.Empty<BorderState>();
            insideVertical = default;
            insideHorizontal = default;
            if (fullRange == null)
            {
                return;
            }

            int cellCount = fullRange.Cells.Count;
            left = new BorderState[cellCount];
            right = new BorderState[cellCount];
            top = new BorderState[cellCount];
            bottom = new BorderState[cellCount];

            int index = 0;
            foreach (Excel.Range cell in fullRange.Cells)
            {
                left[index] = BorderState.From(cell.Borders[Excel.XlBordersIndex.xlEdgeLeft]);
                right[index] = BorderState.From(cell.Borders[Excel.XlBordersIndex.xlEdgeRight]);
                top[index] = BorderState.From(cell.Borders[Excel.XlBordersIndex.xlEdgeTop]);
                bottom[index] = BorderState.From(cell.Borders[Excel.XlBordersIndex.xlEdgeBottom]);
                index++;
            }

            insideVertical = BorderState.From(fullRange.Borders[Excel.XlBordersIndex.xlInsideVertical]);
            insideHorizontal = BorderState.From(fullRange.Borders[Excel.XlBordersIndex.xlInsideHorizontal]);
        }

        private static void RestoreBorders(
            Excel.Range fullRange,
            BorderState[] left,
            BorderState[] right,
            BorderState[] top,
            BorderState[] bottom,
            BorderState insideVertical,
            BorderState insideHorizontal)
        {
            if (fullRange == null || left == null || right == null || top == null || bottom == null)
            {
                return;
            }

            int index = 0;
            foreach (Excel.Range cell in fullRange.Cells)
            {
                left[index].Apply(cell.Borders[Excel.XlBordersIndex.xlEdgeLeft]);
                right[index].Apply(cell.Borders[Excel.XlBordersIndex.xlEdgeRight]);
                top[index].Apply(cell.Borders[Excel.XlBordersIndex.xlEdgeTop]);
                bottom[index].Apply(cell.Borders[Excel.XlBordersIndex.xlEdgeBottom]);
                index++;
            }

            insideVertical.Apply(fullRange.Borders[Excel.XlBordersIndex.xlInsideVertical]);
            insideHorizontal.Apply(fullRange.Borders[Excel.XlBordersIndex.xlInsideHorizontal]);
        }

        private static void RestoreLeftRightAndInside(
            Excel.Range fullRange,
            BorderState[] left,
            BorderState[] right,
            BorderState insideVertical,
            BorderState insideHorizontal)
        {
            if (fullRange == null || left == null || right == null)
            {
                return;
            }

            int index = 0;
            foreach (Excel.Range cell in fullRange.Cells)
            {
                left[index].Apply(cell.Borders[Excel.XlBordersIndex.xlEdgeLeft]);
                right[index].Apply(cell.Borders[Excel.XlBordersIndex.xlEdgeRight]);
                index++;
            }

            try
            {
                insideVertical.Apply(fullRange.Borders[Excel.XlBordersIndex.xlInsideVertical]);
                insideHorizontal.Apply(fullRange.Borders[Excel.XlBordersIndex.xlInsideHorizontal]);
            }
            catch
            {
                // ignored
            }
        }

        private static void ApplySourceFormatting(Excel.Range sourceCell, Excel.Range targetRange)
        {
            if (sourceCell == null || targetRange == null)
            {
                return;
            }

            targetRange.Font.Name = sourceCell.Font.Name;
            targetRange.Font.Size = sourceCell.Font.Size;
            targetRange.Font.Bold = sourceCell.Font.Bold;
            targetRange.Font.Italic = sourceCell.Font.Italic;
            targetRange.Font.Underline = sourceCell.Font.Underline;
            targetRange.Font.Strikethrough = sourceCell.Font.Strikethrough;
            targetRange.Font.Color = sourceCell.Font.Color;
            targetRange.NumberFormat = sourceCell.NumberFormat;
            if (Convert.ToInt32(sourceCell.Interior.ColorIndex) == (int)Excel.XlColorIndex.xlColorIndexNone)
            {
                targetRange.Interior.Pattern = Excel.XlPattern.xlPatternNone;
            }
            else
            {
                targetRange.Interior.Color = sourceCell.Interior.Color;
            }

            targetRange.HorizontalAlignment = sourceCell.HorizontalAlignment;
            targetRange.VerticalAlignment = sourceCell.VerticalAlignment;
            targetRange.WrapText = sourceCell.WrapText;
            targetRange.Orientation = sourceCell.Orientation;
            targetRange.AddIndent = sourceCell.AddIndent;
            targetRange.IndentLevel = sourceCell.IndentLevel;
            targetRange.ShrinkToFit = sourceCell.ShrinkToFit;
        }

        private static void ApplyTopBottomBorders(Excel.Range sourceCell, Excel.Range targetRange)
        {
            if (sourceCell == null || targetRange == null)
            {
                return;
            }

            var top = BorderState.From(sourceCell.Borders[Excel.XlBordersIndex.xlEdgeTop]);
            var bottom = BorderState.From(sourceCell.Borders[Excel.XlBordersIndex.xlEdgeBottom]);

            bool applyTop = top.HasLine;
            bool applyBottom = bottom.HasLine;
            if (!applyTop && !applyBottom)
            {
                return;
            }

            foreach (Excel.Range cell in targetRange.Cells)
            {
                if (applyTop)
                {
                    top.Apply(cell.Borders[Excel.XlBordersIndex.xlEdgeTop]);
                }
                if (applyBottom)
                {
                    bottom.Apply(cell.Borders[Excel.XlBordersIndex.xlEdgeBottom]);
                }
            }
        }

        private void ApplyFinanceRowFormatting(Excel.Range rowSpan, Excel.Range targetRange, Excel.Range sourceCell)
        {
            if (rowSpan == null || sourceCell == null)
            {
                return;
            }

            if (targetRange != null)
            {
                const string format = "$#,##0_);($#,##0);$--_)";
                try
                {
                    targetRange.NumberFormat = format;
                }
                catch
                {
                    try
                    {
                        targetRange.NumberFormatLocal = format;
                    }
                    catch
                    {
                        // ignore
                    }
                }

                targetRange.Font.Name = "Garamond";
                targetRange.Font.Bold = true;

                if (Convert.ToInt32(sourceCell.Interior.ColorIndex) != (int)Excel.XlColorIndex.xlColorIndexNone)
                {
                    targetRange.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    targetRange.Interior.Color = sourceCell.Interior.Color;
                }
                else
                {
                    targetRange.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                }
            }

            rowSpan.Font.Name = "Garamond";
            rowSpan.Font.Bold = true;

            var top = rowSpan.Borders[Excel.XlBordersIndex.xlEdgeTop];
            top.LineStyle = Excel.XlLineStyle.xlContinuous;
            top.Weight = Excel.XlBorderWeight.xlThin;
            top.Color = ColorTranslator.ToOle(Color.Black);
        }

        private bool TryGetBorderDescriptor(string key, out BorderDescriptor descriptor)
        {
            descriptor = null;
            if (string.IsNullOrWhiteSpace(key))
            {
                return false;
            }

            return BorderDescriptors.TryGetValue(key, out descriptor);
        }

        private void ToggleBorder(BorderDescriptor descriptor, int lineStyle, int weight)
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var selection))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                bool state = GetNextBorderState(descriptor.Key);
                var style = NormalizeLineStyle(lineStyle);
                var borderWeight = NormalizeBorderWeight(weight);
                var color = BorderColorSpec.Automatic;
                var mode = state ? BorderOperation.Set : BorderOperation.Delete;
                ApplyBorders(selection, descriptor.Indexes, mode, style, borderWeight, color);
            }
        }

        private void DeleteBorder(BorderDescriptor descriptor)
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var selection))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                ApplyBorders(selection, descriptor.Indexes, BorderOperation.Delete, Excel.XlLineStyle.xlLineStyleNone, Excel.XlBorderWeight.xlThin, BorderColorSpec.Automatic);
                ResetBorderCycleKeys(descriptor.RelatedKeys);
            }
        }

        private void ApplyBorderColor(BorderDescriptor descriptor, BorderColorSpec spec)
        {
            if (!RangeHelpers.TryGetActiveRange(_app, out var selection))
            {
                return;
            }

            using (new UiGuard(_app))
            {
                ApplyBorders(selection, descriptor.Indexes, BorderOperation.ColorOnly, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, spec);
            }
        }

        private Excel.XlLineStyle NormalizeLineStyle(int style)
        {
            try
            {
                var resolved = (Excel.XlLineStyle)style;
                if (!Enum.IsDefined(typeof(Excel.XlLineStyle), resolved) || resolved == Excel.XlLineStyle.xlLineStyleNone)
                {
                    return Excel.XlLineStyle.xlContinuous;
                }

                return resolved;
            }
            catch
            {
                return Excel.XlLineStyle.xlContinuous;
            }
        }

        private Excel.XlBorderWeight NormalizeBorderWeight(int weight)
        {
            try
            {
                var resolved = (Excel.XlBorderWeight)weight;
                if (!Enum.IsDefined(typeof(Excel.XlBorderWeight), resolved))
                {
                    return Excel.XlBorderWeight.xlThin;
                }

                return resolved;
            }
            catch
            {
                return Excel.XlBorderWeight.xlThin;
            }
        }

        private bool GetNextBorderState(string key)
        {
            EnsureBorderCycleFresh();
            bool next = true;
            if (_borderCycleStates.TryGetValue(key, out var current))
            {
                next = !current;
            }

            _borderCycleStates[key] = next;
            return next;
        }

        private void ResetBorderCycleKeys(IEnumerable<string> keys)
        {
            if (keys == null)
            {
                return;
            }

            EnsureBorderCycleFresh();
            foreach (var key in keys)
            {
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                _borderCycleStates.Remove(key);
            }
        }

        private void EnsureBorderCycleFresh()
        {
            if (_borderCycleStampSeen == _selectionStamp)
            {
                return;
            }

            _borderCycleStampSeen = _selectionStamp;
            _borderCycleStates.Clear();
        }

        private void ApplyBorders(
            Excel.Range selection,
            Excel.XlBordersIndex[] indexes,
            BorderOperation operation,
            Excel.XlLineStyle style,
            Excel.XlBorderWeight weight,
            BorderColorSpec colorSpec)
        {
            if (!RangeHelpers.IsRangeValid(selection) || indexes == null || indexes.Length == 0)
            {
                return;
            }

            foreach (var index in indexes)
            {
                if (!ShouldApplyIndex(selection, index))
                {
                    continue;
                }

                Excel.Border border = null;
                try
                {
                    border = selection.Borders[index];
                    if (border == null)
                    {
                        continue;
                    }

                    switch (operation)
                    {
                        case BorderOperation.Delete:
                            border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            break;
                        case BorderOperation.Set:
                            border.LineStyle = style;
                            if ((Excel.XlLineStyle)Convert.ToInt32(border.LineStyle) != Excel.XlLineStyle.xlLineStyleNone)
                            {
                                border.Weight = weight;
                                ApplyBorderColor(border, colorSpec, allowAutomatic: true);
                            }
                            break;
                        case BorderOperation.ColorOnly:
                            var currentStyle = (Excel.XlLineStyle)Convert.ToInt32(border.LineStyle);
                            if (currentStyle != Excel.XlLineStyle.xlLineStyleNone)
                            {
                                ApplyBorderColor(border, colorSpec, allowAutomatic: false);
                            }
                            break;
                    }
                }
                catch
                {
                    // ignore failures on a specific edge
                }
                finally
                {
                    ReleaseIfNeeded(border);
                }
            }
        }

        private bool ShouldApplyIndex(Excel.Range selection, Excel.XlBordersIndex index)
        {
            if (index == Excel.XlBordersIndex.xlInsideHorizontal && selection.Rows.Count == 1)
            {
                return false;
            }

            if (index == Excel.XlBordersIndex.xlInsideVertical && selection.Columns.Count == 1)
            {
                return false;
            }

            return true;
        }

        private void ApplyBorderColor(Excel.Border border, BorderColorSpec colorSpec, bool allowAutomatic)
        {
            if (border == null)
            {
                return;
            }

            try
            {
                if (colorSpec.IsAutomatic)
                {
                    if (allowAutomatic)
                    {
                        border.ColorIndex = (int)Excel.XlColorIndex.xlColorIndexAutomatic;
                    }
                    return;
                }

                if (colorSpec.IsThemeColor)
                {
                    border.ThemeColor = (Excel.XlThemeColor)colorSpec.ThemeColor;
                    border.TintAndShade = colorSpec.TintAndShade;
                }
                else
                {
                    border.Color = colorSpec.Rgb;
                }
            }
            catch
            {
                // ignore color failures
            }
        }

        private int ComputeLastColFromNearestRow(Excel.Worksheet ws, int baseRow, int startCol, int maxOffset, int skipRowStart, int skipRowEnd)
        {
            for (int offset = 1; offset <= maxOffset; offset++)
            {
                int last;
                int upRow = baseRow - offset;
                if (upRow >= 1 && !IsRowSkipped(upRow, skipRowStart, skipRowEnd))
                {
                    last = ContiguousSpanLastCol(ws, upRow, startCol, ignoreBorders: true);
                    if (last > startCol)
                    {
                        return last;
                    }
                }

                int downRow = baseRow + offset;
                if (downRow <= ws.Rows.Count && !IsRowSkipped(downRow, skipRowStart, skipRowEnd))
                {
                    last = ContiguousSpanLastCol(ws, downRow, startCol, ignoreBorders: true);
                    if (last > startCol)
                    {
                        return last;
                    }
                }
            }

            return startCol;
        }

        private static bool IsRowSkipped(int row, int skipStart, int skipEnd)
        {
            if (skipStart <= 0 || skipEnd <= 0)
            {
                return false;
            }

            if (skipStart > skipEnd)
            {
                var tmp = skipStart;
                skipStart = skipEnd;
                skipEnd = tmp;
            }

            return row >= skipStart && row <= skipEnd;
        }

        private int ContiguousSpanLastCol(Excel.Worksheet ws, int rowIndex, int startCol, bool ignoreBorders)
        {
            int last = startCol;
            int maxCol = ws.Columns.Count;

            for (int col = startCol + 1; col <= maxCol; col++)
            {
                var cell = ws.Cells[rowIndex, col] as Excel.Range;
                if (cell == null)
                {
                    break;
                }

                bool hasVisual = HasAnyVisual(cell, ignoreBorders);
                ReleaseIfNeeded(cell);
                if (hasVisual)
                {
                    last = col;
                }
                else
                {
                    break;
                }
            }

            return last;
        }

        private static bool HasAnyVisual(Excel.Range cell)
            => HasAnyVisual(cell, includeBorders: true);

        private static bool HasAnyVisual(Excel.Range cell, bool includeBorders)
        {
            if (cell == null)
            {
                return false;
            }

            try
            {
                var value = cell.Value2;
                bool hasVal = value != null && !string.IsNullOrWhiteSpace(Convert.ToString(value));

                bool hasFill = Convert.ToInt32(cell.DisplayFormat.Interior.ColorIndex) != (int)Excel.XlColorIndex.xlColorIndexNone ||
                               Convert.ToInt32(cell.Interior.ColorIndex) != (int)Excel.XlColorIndex.xlColorIndexNone;

                bool hasBorder = false;
                if (includeBorders)
                {
                    hasBorder = Convert.ToInt32(cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle) != (int)Excel.XlLineStyle.xlLineStyleNone ||
                                Convert.ToInt32(cell.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle) != (int)Excel.XlLineStyle.xlLineStyleNone ||
                                Convert.ToInt32(cell.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle) != (int)Excel.XlLineStyle.xlLineStyleNone ||
                                Convert.ToInt32(cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle) != (int)Excel.XlLineStyle.xlLineStyleNone ||
                                Convert.ToInt32(cell.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle) != (int)Excel.XlLineStyle.xlLineStyleNone ||
                                Convert.ToInt32(cell.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle) != (int)Excel.XlLineStyle.xlLineStyleNone;
                }

                return hasVal || hasFill || hasBorder;
            }
            catch
            {
                return false;
            }
        }

        private void HighlightOutlineCell(
            Excel.Worksheet ws,
            int rowIdx,
            int colIdx,
            int highlightColor,
            int firstRow,
            int lastRow,
            int firstCol,
            int lastCol,
            bool skipTop,
            bool skipLeft,
            HashSet<string> processed)
        {
            if (ws == null || processed == null)
            {
                return;
            }

            if (rowIdx < 1 || colIdx < 1)
            {
                return;
            }

            int maxRow = Convert.ToInt32(ws.Rows.Count);
            int maxCol = Convert.ToInt32(ws.Columns.Count);
            if (rowIdx > maxRow || colIdx > maxCol)
            {
                return;
            }

            string key = $"{rowIdx}|{colIdx}";
            if (!processed.Add(key))
            {
                return;
            }

            Excel.Range cell = null;
            try
            {
                cell = ws.Cells[rowIdx, colIdx] as Excel.Range;
                if (cell == null)
                {
                    return;
                }

                var value = cell.Value2;
                if (value != null && !string.IsNullOrWhiteSpace(Convert.ToString(value)))
                {
                    return;
                }

                cell.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                cell.Interior.Color = highlightColor;

                bool isCorner = false;
                if (rowIdx == firstRow && colIdx == firstCol)
                {
                    isCorner = !skipTop && !skipLeft;
                }
                else if (rowIdx == firstRow && colIdx == lastCol)
                {
                    isCorner = true;
                }
                else if (rowIdx == lastRow && colIdx == firstCol)
                {
                    isCorner = true;
                }
                else if (rowIdx == lastRow && colIdx == lastCol)
                {
                    isCorner = true;
                }

                if (isCorner)
                {
                    cell.Value2 = "x";
                    cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    cell.Font.Name = "Garamond";
                    cell.Font.Size = 11;
                    cell.Font.Color = (int)Excel.XlRgbColor.rgbWhite;
                    cell.Font.Bold = true;
                }
            }
            finally
            {
                ReleaseIfNeeded(cell);
            }
        }

        private void ApplyDownFillFormatting(Excel.Range targetRange, Excel.Range sourceCell, bool sourceHasFill, int sourceFillColor)
        {
            if (targetRange == null || sourceCell == null)
            {
                return;
            }

            targetRange.Font.Name = sourceCell.Font.Name;
            targetRange.Font.Size = sourceCell.Font.Size;
            targetRange.Font.Bold = sourceCell.Font.Bold;
            targetRange.Font.Italic = sourceCell.Font.Italic;
            targetRange.NumberFormat = sourceCell.NumberFormat;

            if (sourceHasFill)
            {
                targetRange.Interior.Color = sourceFillColor;
            }
            else
            {
                targetRange.Interior.Pattern = Excel.XlPattern.xlPatternNone;
            }
        }

        private bool ColumnHasData(Excel.Worksheet ws, int columnIndex, int startRow, int maxRow)
        {
            if (ws == null || startRow > maxRow)
            {
                return false;
            }

            Excel.Range slice = null;
            try
            {
                slice = ws.Range[ws.Cells[startRow, columnIndex], ws.Cells[maxRow, columnIndex]];
                double count = Convert.ToDouble(_app.WorksheetFunction.CountA(slice));
                return count > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseIfNeeded(slice);
            }
        }

        private int GetLastDataRowOrStart(Excel.Worksheet ws, int columnIndex, int startRow, int maxRow, int fallbackRow)
        {
            if (!ColumnHasData(ws, columnIndex, startRow, maxRow))
            {
                return fallbackRow;
            }

            Excel.Range lastCell = null;
            Excel.Range bottomCell = null;
            try
            {
                bottomCell = ws.Cells[maxRow, columnIndex] as Excel.Range;
                if (bottomCell == null)
                {
                    return fallbackRow;
                }

                lastCell = bottomCell.End[Excel.XlDirection.xlUp] as Excel.Range;
                if (lastCell == null)
                {
                    return fallbackRow;
                }

                int row = lastCell.Row;
                return Math.Max(row, fallbackRow);
            }
            catch
            {
                return fallbackRow;
            }
            finally
            {
                ReleaseIfNeeded(bottomCell);
                ReleaseIfNeeded(lastCell);
            }
        }

        private static bool HasCellValue(Excel.Range cell)
        {
            if (cell == null)
            {
                return false;
            }

            try
            {
                var value = cell.Value2;
                if (value == null)
                {
                    return false;
                }

                if (value is double || value is float || value is decimal)
                {
                    return true;
                }

                if (value is DateTime)
                {
                    return true;
                }

                var text = Convert.ToString(value);
                return !string.IsNullOrWhiteSpace(text);
            }
            catch
            {
                return false;
            }
        }

        private Excel.Range MergeIntoUnion(Excel.Range unionRange, Excel.Range addition)
        {
            if (addition == null)
            {
                return unionRange;
            }

            if (unionRange == null)
            {
                return addition;
            }

            Excel.Range merged;
            try
            {
                merged = _app.Union(unionRange, addition);
            }
            catch
            {
                ReleaseIfNeeded(addition);
                return unionRange;
            }

            ReleaseIfNeeded(unionRange);
            ReleaseIfNeeded(addition);
            return merged;
        }

        private static string WrapCircFormulaText(string formula)
        {
            if (string.IsNullOrEmpty(formula) || formula[0] != '=')
            {
                return formula;
            }

            var inner = formula.Substring(1);
            return $"=IF(circ=1,0,{inner})";
        }

        private void SelectRangeSafe(Excel.Range range)
        {
            if (!RangeHelpers.IsRangeValid(range))
            {
                return;
            }

            RangeHelpers.SafeSelect(range);
        }

        private readonly struct BorderState
        {
            private readonly Excel.XlLineStyle _lineStyle;
            private readonly Excel.XlBorderWeight _weight;
            private readonly int _color;

            private BorderState(Excel.XlLineStyle style, Excel.XlBorderWeight weight, int color)
            {
                _lineStyle = style;
                _weight = weight;
                _color = color;
            }

            public static BorderState From(Excel.Border border)
            {
                var style = (Excel.XlLineStyle)Convert.ToInt32(border.LineStyle);
                var weight = (Excel.XlBorderWeight)Convert.ToInt32(border.Weight);
                int color = Convert.ToInt32(border.Color);
                return new BorderState(style, weight, color);
            }

            public bool HasLine => _lineStyle != Excel.XlLineStyle.xlLineStyleNone;

            public void Apply(Excel.Border border)
            {
                border.LineStyle = _lineStyle;
                if (_lineStyle != Excel.XlLineStyle.xlLineStyleNone)
                {
                    border.Weight = _weight;
                    border.Color = _color;
                }
                else
                {
                    border.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                }
            }
        }
        #endregion

        #region CopySelection helpers
        private bool CopySelectionAsPicture(object target)
        {
            if (target == null)
            {
                return false;
            }

            if (target is Excel.Range range)
            {
                return CopyRangePictureReliable(range);
            }

            if (TryCopyAsPicture(target, Excel.XlPictureAppearance.xlPrinter, Excel.XlCopyPictureFormat.xlPicture))
            {
                return true;
            }

            if (TryCopyAsPicture(target, Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture))
            {
                return true;
            }

            object parent = target;
            for (int i = 0; i < 10 && parent != null; i++)
            {
                parent = GetParent(parent);
                if (parent == null)
                {
                    break;
                }

                if (TryCopyAsPicture(parent, Excel.XlPictureAppearance.xlPrinter, Excel.XlCopyPictureFormat.xlPicture))
                {
                    return true;
                }
                if (TryCopyAsPicture(parent, Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture))
                {
                    return true;
                }
            }

            return false;
        }

        private bool CopyRangePictureReliable(Excel.Range rng)
        {
            if (!RangeHelpers.IsRangeValid(rng))
            {
                return false;
            }

            Excel.Worksheet ws = rng.Worksheet;
            Excel.Range area = rng.Areas.Count > 1 ? rng.Areas[1] : rng;

            Excel.Window originalWindow = null;
            Excel.Worksheet originalSheet = null;

            try
            {
                originalWindow = _app.ActiveWindow;
                originalSheet = originalWindow?.ActiveSheet as Excel.Worksheet;
            }
            catch
            {
                // ignore
            }

            for (int attempt = 1; attempt <= 4; attempt++)
            {
                _app.CutCopyMode = Excel.XlCutCopyMode.xlCopy;

                try
                {
                    if (ws != null && ws != _app.ActiveSheet)
                    {
                        ws.Activate();
                    }
                }
                catch
                {
                    // ignore
                }

                var anchorCell = area.Cells[1, 1] as Excel.Range;
                if (anchorCell != null)
                {
                    RangeHelpers.SafeSelect(anchorCell);
                }
                try
                {
                    switch (attempt)
                    {
                        case 1:
                            area.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);
                            break;
                        case 2:
                            area.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);
                            break;
                        case 3:
                            area.CopyPicture(Excel.XlPictureAppearance.xlPrinter, Excel.XlCopyPictureFormat.xlPicture);
                            break;
                        default:
                            area.CopyPicture(Excel.XlPictureAppearance.xlPrinter, Excel.XlCopyPictureFormat.xlBitmap);
                            break;
                    }

                    WaitForClipboardReady(400);
                    Excel.ShapeRange pastedRange = null;
                    try
                    {
                        dynamic shapes = ws.Shapes;
                        pastedRange = shapes.Paste();
                    }
                    catch
                    {
                        pastedRange = null;
                    }
                    if (pastedRange == null || pastedRange.Count == 0)
                    {
                        continue;
                    }

                    var shape = pastedRange.Item(1) as Excel.Shape;
                    if (shape == null)
                    {
                        continue;
                    }

                    double w = Convert.ToDouble(area.Width);
                    double h = Convert.ToDouble(area.Height);
                    bool okSize = Math.Abs(Convert.ToDouble(shape.Width) - w) <= (w * 0.1 + 2) &&
                                  Math.Abs(Convert.ToDouble(shape.Height) - h) <= (h * 0.1 + 2);
                    shape.Copy();
                    WaitForClipboardReady(500);
                    shape.Delete();
                    if (okSize)
                    {
                        return true;
                    }
                }
                catch
                {
                    // ignore and retry
                }

                System.Windows.Forms.Application.DoEvents();
            }

            try
            {
                originalWindow?.Activate();
                originalSheet?.Activate();
            }
            catch
            {
                // ignore
            }

            return false;
        }

        private bool TryCopyAsPicture(object target, Excel.XlPictureAppearance appearance, Excel.XlCopyPictureFormat format)
        {
            try
            {
                switch (target)
                {
                    case Excel.Range range:
                        range.CopyPicture(appearance, format);
                        break;
                    case Excel.Shape shape:
                        shape.CopyPicture(appearance, format);
                        break;
                    case Excel.Chart chart:
                        chart.CopyPicture(appearance, format);
                        break;
                    default:
                        return false;
                }

                WaitForClipboardReady(400);
                return _clipboard.ClipboardHasContent();
            }
            catch
            {
                return false;
            }
        }

        private object ResolveFallbackCopyTarget(object selection)
        {
            object probe = selection;
            while (probe != null)
            {
                switch (probe)
                {
                    case Excel.Range _:
                    case Excel.ChartObject _:
                    case Excel.Chart _:
                    case Excel.Shape _:
                        return probe;
                }

                probe = GetParent(probe);
            }

            try
            {
                if (_app.ActiveChart != null)
                {
                    return _app.ActiveChart;
                }
            }
            catch
            {
                // ignore
            }

            if (_app.Selection is Excel.Range range)
            {
                return range;
            }

            return null;
        }

        private static object GetParent(object obj)
        {
            try
            {
                return obj.GetType().GetProperty("Parent")?.GetValue(obj);
            }
            catch
            {
                return null;
            }
        }

        private void WaitForClipboardReady(int maxMillis)
        {
            var start = DateTime.UtcNow;
            while ((DateTime.UtcNow - start).TotalMilliseconds < maxMillis)
            {
                if (_clipboard.ClipboardHasContent())
                {
                    return;
                }

                System.Windows.Forms.Application.DoEvents();
            }
        }
        #endregion

        private sealed class NumberFormatCycleState
        {
            public string LastSelectionKey = string.Empty;
            public int NextIndex;
        }

        private bool TryApplyShapeFill(ColorSpec spec)
        {
            object selection;
            try { selection = _app.Selection; }
            catch { return false; }

            if (selection == null)
            {
                return false;
            }

            bool applied = false;
            switch (selection)
            {
                case Excel.ShapeRange range:
                    foreach (Excel.Shape shape in range)
                    {
                        try { applied |= ApplyFillToShape(shape, spec); }
                        finally { ReleaseIfNeeded(shape); }
                    }
                    ReleaseIfNeeded(range);
                    break;
                case Excel.Shape shape:
                    applied = ApplyFillToShape(shape, spec);
                    ReleaseIfNeeded(shape);
                    break;
                case Excel.ChartObject chartObject:
                    applied = ApplyFillFormat(chartObject.Chart?.ChartArea?.Format?.Fill, spec);
                    ReleaseIfNeeded(chartObject);
                    break;
                case Excel.Chart chart:
                    applied = ApplyFillFormat(chart.ChartArea?.Format?.Fill, spec);
                    ReleaseIfNeeded(chart);
                    break;
                default:
                    break;
            }

            return applied;
        }

        private bool TryApplyShapeFont(ColorSpec spec)
        {
            object selection;
            try { selection = _app.Selection; }
            catch { return false; }

            if (selection == null)
            {
                return false;
            }

            bool applied = false;
            switch (selection)
            {
                case Excel.ShapeRange range:
                    foreach (Excel.Shape shape in range)
                    {
                        try { applied |= ApplyFontToShape(shape, spec); }
                        finally { ReleaseIfNeeded(shape); }
                    }
                    ReleaseIfNeeded(range);
                    break;
                case Excel.Shape shape:
                    applied = ApplyFontToShape(shape, spec);
                    ReleaseIfNeeded(shape);
                    break;
                default:
                    break;
            }

            return applied;
        }

        private bool TryApplyShapeBorder(ColorSpec spec)
        {
            object selection;
            try { selection = _app.Selection; }
            catch { return false; }

            if (selection == null)
            {
                return false;
            }

            bool applied = false;
            switch (selection)
            {
                case Excel.ShapeRange range:
                    foreach (Excel.Shape shape in range)
                    {
                        try { applied |= ApplyBorderToShape(shape, spec); }
                        finally { ReleaseIfNeeded(shape); }
                    }
                    ReleaseIfNeeded(range);
                    break;
                case Excel.Shape shape:
                    applied = ApplyBorderToShape(shape, spec);
                    ReleaseIfNeeded(shape);
                    break;
                default:
                    break;
            }

            return applied;
        }

        private void ApplyInterior(Excel.Interior interior, ColorSpec spec)
        {
            if (interior == null)
            {
                return;
            }

            try
            {
                if (spec.IsNull)
                {
                    interior.Pattern = Excel.XlPattern.xlPatternNone;
                    interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                }
                else if (spec.IsThemeColor)
                {
                    interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    try
                    {
                        interior.ThemeColor = spec.ThemeColor;
                        interior.TintAndShade = spec.TintAndShade;
                    }
                    catch
                    {
                        interior.Color = spec.Rgb;
                    }
                }
                else
                {
                    interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    interior.Color = spec.Rgb;
                }
            }
            catch
            {
                // ignore
            }
        }

        private bool TryApplyFontColorToSelection(ColorSpec spec)
        {
            object selection;
            try { selection = _app.Selection; }
            catch { return false; }

            if (selection == null)
            {
                return false;
            }

            if (TrySetFontColor(selection, spec))
            {
                return true;
            }

            bool applied = false;
            switch (selection)
            {
                case Excel.DataLabels labels:
                    foreach (Excel.DataLabel label in labels)
                    {
                        try { applied |= TrySetFontColor(label, spec); }
                        finally { ReleaseIfNeeded(label); }
                    }
                    ReleaseIfNeeded(labels);
                    break;
                case Excel.DataLabel label:
                    applied = TrySetFontColor(label, spec);
                    ReleaseIfNeeded(label);
                    break;
                case Excel.Point dataPoint:
                    if (dataPoint.HasDataLabel)
                    {
                        applied = TrySetFontColor(dataPoint.DataLabel, spec);
                    }
                    ReleaseIfNeeded(dataPoint);
                    break;
                case Excel.Series series:
                    if (series.HasDataLabels)
                    {
                        Excel.DataLabels seriesLabels = null;
                        try
                        {
                            seriesLabels = (Excel.DataLabels)series.DataLabels();
                        }
                        catch
                        {
                            seriesLabels = null;
                        }
                        if (seriesLabels != null)
                        {
                            int labelCount = seriesLabels.Count;
                            for (int idx = 1; idx <= labelCount; idx++)
                            {
                                Excel.DataLabel dl = null;
                                try
                                {
                                    dl = seriesLabels.Item(idx) as Excel.DataLabel;
                                    if (dl == null)
                                    {
                                        continue;
                                    }
                                    applied |= TrySetFontColor(dl, spec);
                                }
                                finally { ReleaseIfNeeded(dl); }
                            }
                            ReleaseIfNeeded(seriesLabels);
                        }
                    }
                    ReleaseIfNeeded(series);
                    break;
                case Excel.ShapeRange shapeRange:
                    foreach (Excel.Shape shape in shapeRange)
                    {
                        try { applied |= ApplyFontToShape(shape, spec); }
                        finally { ReleaseIfNeeded(shape); }
                    }
                    ReleaseIfNeeded(shapeRange);
                    break;
                case Excel.Chart chart:
                    applied |= ApplyChartFonts(chart, spec);
                    ReleaseIfNeeded(chart);
                    break;
                default:
                    break;
            }

            return applied;
        }

        private bool ApplyChartFonts(Excel.Chart chart, ColorSpec spec)
        {
            bool applied = false;
            try
            {
                if (chart.HasTitle)
                {
                    applied |= TrySetFontColor(chart.ChartTitle, spec);
                }

                if (chart.HasDataTable)
                {
                    applied |= TrySetFontColor(chart.DataTable, spec);
                }

                if (chart.HasLegend)
                {
                    applied |= TrySetFontColor(chart.Legend, spec);
                }
            }
            catch
            {
                // ignore
            }

            return applied;
        }

        private bool TrySetFontColor(object target, ColorSpec spec)
        {
            if (target == null)
            {
                return false;
            }

            try
            {
                dynamic dyn = target;
                var font = dyn.Font;
                if (font == null)
                {
                    return false;
                }

                ApplyFontColor(font, spec);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void ApplyFontColor(dynamic font, ColorSpec spec)
        {
            if (font == null)
            {
                return;
            }

            try
            {
                if (spec.IsNull)
                {
                    font.ColorIndex = (int)Excel.XlColorIndex.xlColorIndexAutomatic;
                }
                else if (spec.IsThemeColor)
                {
                    try
                    {
                        font.ThemeColor = spec.ThemeColor;
                        font.TintAndShade = spec.TintAndShade;
                    }
                    catch
                    {
                        font.Color = spec.Rgb;
                    }
                }
                else
                {
                    font.Color = spec.Rgb;
                }
            }
            catch
            {
                // ignore
            }
        }

        private bool ApplyFillToShape(Excel.Shape shape, ColorSpec spec)
        {
            if (shape == null)
            {
                return false;
            }

            try
            {
                if (shape.Fill == null)
                {
                    return false;
                }

                return ApplyFillFormat(shape.Fill, spec);
            }
            catch
            {
                return false;
            }
        }

        private bool ApplyFontToShape(Excel.Shape shape, ColorSpec spec)
        {
            if (shape == null)
            {
                return false;
            }

            bool applied = false;
            try
            {
                if (shape.TextFrame2 != null && shape.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                {
                    var fill = shape.TextFrame2.TextRange.Font.Fill;
                    applied |= ApplyThemeFill(fill, spec);
                }

                if (shape.HasChart == Office.MsoTriState.msoTrue)
                {
                    applied |= ApplyChartFonts(shape.Chart, spec);
                }
            }
            catch
            {
                // ignore
            }

            return applied;
        }

        private bool ApplyBorderToShape(Excel.Shape shape, ColorSpec spec)
        {
            if (shape == null)
            {
                return false;
            }

            try
            {
                var line = shape.Line;
                if (line == null)
                {
                    return false;
                }

                if (spec.IsNull)
                {
                    line.Visible = Office.MsoTriState.msoFalse;
                }
                else
                {
                    line.Visible = Office.MsoTriState.msoTrue;
                    if (spec.IsThemeColor)
                    {
                        try
                        {
                            line.ForeColor.ObjectThemeColor = (Office.MsoThemeColorIndex)spec.ObjectThemeColor;
                            line.ForeColor.TintAndShade = (float)spec.TintAndShade;
                        }
                        catch
                        {
                            line.ForeColor.RGB = spec.Rgb;
                        }
                    }
                    else
                    {
                        line.ForeColor.RGB = spec.Rgb;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool ApplyFillFormat(object fillObject, ColorSpec spec)
        {
            if (fillObject == null)
            {
                return false;
            }

            try
            {
                dynamic fill = fillObject;
                if (spec.IsNull)
                {
                    fill.Visible = Office.MsoTriState.msoFalse;
                    fill.Transparency = 1f;
                }
                else
                {
                    fill.Visible = Office.MsoTriState.msoTrue;
                    fill.Solid();
                    if (spec.IsThemeColor)
                    {
                        try
                        {
                            fill.ForeColor.ObjectThemeColor = (Office.MsoThemeColorIndex)spec.ObjectThemeColor;
                            fill.ForeColor.TintAndShade = (float)spec.TintAndShade;
                        }
                        catch
                        {
                            fill.ForeColor.RGB = spec.Rgb;
                        }
                    }
                    else
                    {
                        fill.ForeColor.RGB = spec.Rgb;
                    }
                    fill.Transparency = 0f;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool ApplyThemeFill(object fillObject, ColorSpec spec)
        {
            if (fillObject == null)
            {
                return false;
            }

            try
            {
                dynamic fill = fillObject;
                if (spec.IsNull)
                {
                    fill.Visible = Office.MsoTriState.msoFalse;
                }
                else if (spec.IsThemeColor)
                {
                    fill.Visible = Office.MsoTriState.msoTrue;
                    fill.ForeColor.ObjectThemeColor = (Office.MsoThemeColorIndex)spec.ObjectThemeColor;
                    fill.ForeColor.TintAndShade = (float)spec.TintAndShade;
                }
                else
                {
                    fill.Visible = Office.MsoTriState.msoTrue;
                    fill.ForeColor.RGB = spec.Rgb;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private enum BorderOperation
        {
            Set,
            Delete,
            ColorOnly
        }

        private readonly struct BorderColorSpec
        {
            public static BorderColorSpec Automatic => new BorderColorSpec(true, false, 0, 0, 0);

            public BorderColorSpec(bool isAutomatic, bool isThemeColor, int themeColor, double tintAndShade, int rgb)
            {
                IsAutomatic = isAutomatic;
                IsThemeColor = isThemeColor;
                ThemeColor = themeColor;
                TintAndShade = tintAndShade;
                Rgb = rgb;
            }

            public bool IsAutomatic { get; }
            public bool IsThemeColor { get; }
            public int ThemeColor { get; }
            public double TintAndShade { get; }
            public int Rgb { get; }
        }

        private sealed class BorderDescriptor
        {
            public BorderDescriptor(string key, Excel.XlBordersIndex[] indexes, string[] relatedKeys)
            {
                Key = key ?? string.Empty;
                Indexes = indexes ?? Array.Empty<Excel.XlBordersIndex>();
                RelatedKeys = relatedKeys ?? Array.Empty<string>();
            }

            public string Key { get; }
            public Excel.XlBordersIndex[] Indexes { get; }
            public string[] RelatedKeys { get; }
        }

        private readonly struct ColorSpec
        {
            public ColorSpec(bool isNull, bool isThemeColor, int themeColor, int objectThemeColor, double tintAndShade, int rgb)
            {
                IsNull = isNull;
                IsThemeColor = isThemeColor;
                ThemeColor = themeColor;
                ObjectThemeColor = objectThemeColor;
                TintAndShade = tintAndShade;
                Rgb = rgb;
            }

            public bool IsNull { get; }
            public bool IsThemeColor { get; }
            public int ThemeColor { get; }
            public int ObjectThemeColor { get; }
            public double TintAndShade { get; }
            public int Rgb { get; }
        }

        private object GetActiveRange()
        {
            try
            {
                return _app.Selection;
            }
            catch
            {
                return null;
            }
        }

        private void ExecuteMso(string controlId)
        {
            try
            {
                _app.CommandBars.ExecuteMso(controlId);
            }
            catch
            {
                // ignore
            }
        }

        private void RunMacroIfExists(string macroName, params object[] args)
        {
            if (string.IsNullOrWhiteSpace(macroName))
            {
                return;
            }

            try
            {
                if (args == null || args.Length == 0)
                {
                    _app.Run(macroName);
                }
                else
                {
                    _app.Run(macroName, args);
                }
            }
            catch
            {
                // ignore if macro missing
            }
        }

        private static void ReleaseIfNeeded(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comObject);
            }
            catch
            {
                // ignore
            }
        }
    }
}

