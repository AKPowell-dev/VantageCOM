using System;
using System.Runtime.InteropServices;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    [ComVisible(true)]
    [Guid("7D759476-0E72-4B44-B296-FFACDC61CCAA")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class VantageEngine : IDisposable
    {
        static VantageEngine()
        {
            try
            {
                // Early binding check; if dependencies are missing, surface it.
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString(), "VantageEngine static ctor");
            }
        }

        private readonly Excel.Application _excel;
        private readonly Lazy<ClipboardService> _clipboard;
        private readonly Lazy<PowerPointExporter> _ppt;
        private readonly Lazy<FormatService> _format;
        private readonly Lazy<AutoColorService> _autoColor;
        private readonly Lazy<UtilService> _util;
        private readonly Lazy<InsertModeService> _insertMode;
        private readonly Lazy<BatchResizeService> _batchResize;
        private readonly Lazy<WorkbookOptimizer> _optimizer;
        private readonly Lazy<EditingService> _editing;
        private readonly Lazy<WorkbookAnalysisService> _analysis;
        private readonly Lazy<ChartNavigator> _charts;
        private readonly Lazy<TraceDialogService> _traceDialogs;
        private bool _pendingFormatReset;

        public VantageEngine(Excel.Application excel)
        {
            _excel = excel ?? throw new ArgumentNullException(nameof(excel));
            _clipboard = new Lazy<ClipboardService>(() => new ClipboardService(_excel), LazyThreadSafetyMode.None);
            _ppt = new Lazy<PowerPointExporter>(() => new PowerPointExporter(), LazyThreadSafetyMode.None);
            _format = new Lazy<FormatService>(() => new FormatService(_excel, Clipboard, PowerPoint), LazyThreadSafetyMode.None);
            _autoColor = new Lazy<AutoColorService>(() => new AutoColorService(_excel), LazyThreadSafetyMode.None);
            _util = new Lazy<UtilService>(() => new UtilService(_excel), LazyThreadSafetyMode.None);
            _insertMode = new Lazy<InsertModeService>(() => new InsertModeService(_excel), LazyThreadSafetyMode.None);
            _batchResize = new Lazy<BatchResizeService>(() => new BatchResizeService(_excel, Format, PowerPoint), LazyThreadSafetyMode.None);
            _optimizer = new Lazy<WorkbookOptimizer>(() => new WorkbookOptimizer(_excel), LazyThreadSafetyMode.None);
            _editing = new Lazy<EditingService>(() => new EditingService(_excel), LazyThreadSafetyMode.None);
            _analysis = new Lazy<WorkbookAnalysisService>(() => new WorkbookAnalysisService(_excel), LazyThreadSafetyMode.None);
            _charts = new Lazy<ChartNavigator>(() => new ChartNavigator(_excel), LazyThreadSafetyMode.None);
            _traceDialogs = new Lazy<TraceDialogService>(() => new TraceDialogService(_excel), LazyThreadSafetyMode.None);
        }

        public void Dispose() { }

        private ClipboardService Clipboard => _clipboard.Value;
        private PowerPointExporter PowerPoint => _ppt.Value;
        private AutoColorService AutoColor => _autoColor.Value;
        private UtilService Util => _util.Value;
        private InsertModeService InsertMode => _insertMode.Value;
        private BatchResizeService BatchResize => _batchResize.Value;
        private FormatService Format
        {
            get
            {
                var service = _format.Value;
                if (_pendingFormatReset)
                {
                    service.ResetCycleState();
                    _pendingFormatReset = false;
                }

                return service;
            }
        }
        private WorkbookOptimizer Optimizer => _optimizer.Value;
        private EditingService Editing => _editing.Value;
        private WorkbookAnalysisService Analysis => _analysis.Value;
        private ChartNavigator Charts => _charts.Value;
        private TraceDialogService TraceDialogs => _traceDialogs.Value;

        #region Clipboard hooks
        public void ClipboardHandleCopy() => Clipboard.HandleCopy();
        public void ClipboardHandleCut() => Clipboard.HandleCut();
        public void ClipboardHandlePaste() => Clipboard.HandlePaste();
        public void ClipboardHandlePasteValues() => Clipboard.HandlePasteValues();
        public void ClipboardHandlePasteFormulas() => Clipboard.HandlePasteFormulas();
        public void ClipboardOpenPasteSpecial() => Clipboard.OpenPasteSpecialDialog();
        public Excel.Range ClipboardGetCopyRange() => Clipboard.GetCopyRange();
        public void ClipboardSetCopyRange(Excel.Range range) => Clipboard.SetCopyRange(range);
        public void ClipboardPasteValuesSmart() => Clipboard.PasteValuesSmart();
        #endregion

        #region Auto color
        public void AutoColorRange(Excel.Range target, int maxCells) => AutoColor.ApplyAutoColor(target, maxCells);
        #endregion

        #region Utility helpers
        public void TimeClear() => Util.TimeClear();
        public double GetQueryPerformanceTime(string format) => Util.GetQueryPerformanceTime(format);
        public void SetStatusBar(string text, long currentCount, long maximumCount, double percent, int numDigitsAfterDecimal, bool progressBar, bool countPerMax)
            => Util.SetStatusBar(text, currentCount, maximumCount, percent, numDigitsAfterDecimal, progressBar, countPerMax);
        public void SetStatusBarTemporarily(string text, int milliseconds, bool disablePrefix, string statusPrefix)
            => Util.SetStatusBarTemporarily(text, milliseconds, disablePrefix, statusPrefix);
        public bool RegExpMatch(string str, string matchPattern, bool isIgnoreCase, bool isGlobal, bool isMultiline)
            => Util.RegExpMatch(str, matchPattern, isIgnoreCase, isGlobal, isMultiline);
        public string RegExpSearch(string str, string matchPattern, bool isIgnoreCase, bool isGlobal, bool isMultiline)
            => Util.RegExpSearch(str, matchPattern, isIgnoreCase, isGlobal, isMultiline);
        public string RegExpReplace(string str, string matchPattern, string replaceStr, bool isIgnoreCase, bool isGlobal, bool isMultiline)
            => Util.RegExpReplace(str, matchPattern, replaceStr, isIgnoreCase, isGlobal, isMultiline);
        public bool StartsWith(string str, object prefixes) => Util.StartsWith(str, prefixes);
        public bool EndsWith(string str, object suffixes) => Util.EndsWith(str, suffixes);
        public long GetWorkbookIndex(Excel.Workbook targetWorkbook) => Util.GetWorkbookIndex(targetWorkbook);
        public bool IsSheetExists(string targetSheetName) => Util.IsSheetExists(targetSheetName);
        public long GetVisibleSheetsCount() => Util.GetVisibleSheetsCount();
        public object[] DirGrob(string folderPath) => Util.DirGrob(folderPath);
        public string GetAbsolutePath(string cwd, string relativePath) => Util.GetAbsolutePath(cwd, relativePath);
        public string ResolvePath(string strPath) => Util.ResolvePath(strPath);
        public long HexColorCodeToLong(string colorCode) => Util.HexColorCodeToLong(colorCode);
        public string ColorCodeToHex(long colorCode) => Util.ColorCodeToHex(colorCode);
        public bool IsJISKeyboardLayout() => Util.IsJisKeyboardLayout();
        public Excel.Range Union2(object argList) => Util.Union2(argList);
        public Excel.Range Intersect2(object argList) => Util.Intersect2(argList);
        public Excel.Range Except2(object sourceRange, object argList) => Util.Except2(sourceRange, argList);
        public Excel.Range Invert2(object sourceRange) => Util.Invert2(sourceRange);
        public bool IsRangeValid(Excel.Range candidate) => Util.IsRangeValid(candidate);
        public void DebugPrint(string message, string funcName, bool debugMode, string statusPrefix)
            => Util.DebugPrint(message, funcName, debugMode, statusPrefix);
        public bool ErrorHandler(int errNumber, string errDescription, string funcName, string statusPrefix)
            => Util.ErrorHandler(errNumber, errDescription, funcName, statusPrefix);
        #endregion

        #region Insert mode
        public bool InsertWithIME() => InsertMode.InsertWithIME();
        public bool InsertWithoutIME() => InsertMode.InsertWithoutIME();
        public bool AppendWithIME() => InsertMode.AppendWithIME();
        public bool AppendWithoutIME() => InsertMode.AppendWithoutIME();
        public bool SubstituteWithIME() => InsertMode.SubstituteWithIME();
        public bool SubstituteWithoutIME() => InsertMode.SubstituteWithoutIME();
        #endregion

        #region Formatting helpers
        public void PasteExact() => Format.PasteExact();
        public void PasteCondensed() => Format.PasteCondensed();
        public void SmartFillRight() => Format.SmartFillRight();
        public void SmartFormatRight() => Format.SmartFormatRight();
        public void OutlineSelectionHighlight() => Format.OutlineSelectionHighlight();
        public void SmartFillDown() => Format.SmartFillDown();
        public void WrapFormulaWithCircCheck() => Format.WrapFormulaWithCircCheck();
        public void FillFromAbove() => Format.FillFromAbove();
        public void FillFromBelow() => Format.FillFromBelow();
        public void FillFromLeft() => Format.FillFromLeft();
        public void FillFromRight() => Format.FillFromRight();
        public void CopySelectionAsPlainText(string delimiter) => Format.CopySelectionAsPlainText(delimiter);
        public bool CopySelectionAsPicturePrintSafe() => Format.CopySelectionAsPicturePrintSafe();
        public void CopySelectionToPowerPoint() => Format.CopyPasteSelectionToPowerPoint();
        public void FormatChartFg() => Format.FormatChartFg();
        public void IncreaseFontSize(int steps) => Format.IncreaseFontSize(steps);
        public void DecreaseFontSize(int steps) => Format.DecreaseFontSize(steps);
        public void AlignLeft() => Format.AlignLeft();
        public void AlignCenter() => Format.AlignCenter();
        public void AlignRight() => Format.AlignRight();
        public void AlignTop() => Format.AlignTop();
        public void AlignMiddle() => Format.AlignMiddle();
        public void AlignBottom() => Format.AlignBottom();
        public void ToggleBold() => Format.ToggleBold();
        public void ToggleItalic() => Format.ToggleItalic();
        public void ToggleUnderline() => Format.ToggleUnderline();
        public void ToggleStrikethrough() => Format.ToggleStrikethrough();
        public void ShowFontDialog() => Format.ShowFontDialog();
        public void ShowFormatNumberDialog() => Format.ShowFormatNumberDialog();
        public void IncreaseDecimalPlaces(int steps) => Format.IncreaseDecimalPlaces(steps);
        public void DecreaseDecimalPlaces(int steps) => Format.DecreaseDecimalPlaces(steps);
        public void ApplyInteriorColor(bool isNull, bool isTheme, int themeColor, double tint, int rgb) => Format.ApplyInteriorColor(isNull, isTheme, themeColor, tint, rgb);
        public bool ApplyFontColor(bool isNull, bool isTheme, int themeColor, int objectThemeColor, double tint, int rgb) => Format.ApplyFontColor(isNull, isTheme, themeColor, objectThemeColor, tint, rgb);
        public void ApplyShapeFillColor(bool isNull, bool isTheme, int themeColor, double tint, int rgb) => Format.ApplyShapeFillColor(isNull, isTheme, themeColor, tint, rgb);
        public void ApplyShapeFontColor(bool isNull, bool isTheme, int themeColor, int objectThemeColor, double tint, int rgb) => Format.ApplyShapeFontColor(isNull, isTheme, themeColor, objectThemeColor, tint, rgb);
        public void ApplyShapeBorderColor(bool isNull, bool isTheme, int themeColor, double tint, int rgb) => Format.ApplyShapeBorderColor(isNull, isTheme, themeColor, tint, rgb);
        public bool ApplySmartFontColor(bool isNull, bool isTheme, int themeColor, int objectThemeColor, double tint, int rgb) => Format.ApplySmartFontColor(isNull, isTheme, themeColor, objectThemeColor, tint, rgb);
        public void ApplySmartFillColor(bool isNull, bool isTheme, int themeColor, double tint, int rgb) => Format.ApplySmartFillColor(isNull, isTheme, themeColor, tint, rgb);

        public void LockCellReference() => Format.LockCellReference();
        public void ResetCycleState()
        {
            if (_format.IsValueCreated)
            {
                Format.ResetCycleState();
            }
            else
            {
                _pendingFormatReset = true;
            }
        }
        public void ClearFormatting() => Format.ClearFormatting();
        public void CycleFormatting() => Format.CycleFormatting();
        public void CycleNumberFormat(long selectionStamp) => Format.CycleNumberFormat(selectionStamp);
        public void BinaryCycle(long selectionStamp) => Format.BinaryCycle(selectionStamp);
        public void YearDisplayCycle(long selectionStamp) => Format.YearDisplayCycle(selectionStamp);
        public void NumberNarrativeCycle(long selectionStamp) => Format.NumberNarrativeCycle(selectionStamp);
        public void PercentCycle(long selectionStamp) => Format.PercentCycle(selectionStamp);
        public void FlipSign() => Format.FlipSign();
        public void ReverseSelectionOrder() => Format.ReverseSelectionOrder();
        public void TrimConditionalFormatting() => Format.TrimConditionalFormatting();
        public void CurrencyCycle(long selectionStamp) => Format.CurrencyCycle(selectionStamp);
        public void ToggleBorder(string targetKey, int lineStyle, int weight) => Format.ToggleBorder(targetKey, lineStyle, weight);
        public void DeleteBorder(string targetKey) => Format.DeleteBorder(targetKey);
        public void SetBorderColor(string targetKey, bool isNull, bool isTheme, int themeColor, double tintAndShade, int rgb) => Format.SetBorderColor(targetKey, isNull, isTheme, themeColor, tintAndShade, rgb);
        public void ResizeSelectionToWidthInches(double targetInches, bool requirePowerPoint) => BatchResize.ResizeSelectionToWidthInches(targetInches, requirePowerPoint);

        public void AdjustNumbers(int count, bool subtract, bool grow) => Editing.AdjustNumbers(count, subtract, grow);
        public void ApplyAutoFill() => Editing.ApplyAutoFill();
        public void NavigateSpecialCells(int typeValue, Excel.XlSearchOrder searchOrder, bool forward, int steps) => Editing.NavigateSpecialCells(typeValue, searchOrder, forward, steps);
        public void SubstituteType(string text) => Editing.SubstituteType(text);
        public void InsertRows(int count, bool append) => Editing.InsertRows(count, append);
        public void DeleteRows(int targetType, int count) => Editing.DeleteRows(targetType, count);
        public void HideRows(int targetType, int count, bool hide) => Editing.HideRows(targetType, count, hide);
        public void GroupRows(int count, bool group) => Editing.GroupRows(count, group);
        public void InsertColumns(int count, bool append) => Editing.InsertColumns(count, append);
        public void DeleteColumns(int targetType, int count) => Editing.DeleteColumns(targetType, count);
        public void HideColumns(int targetType, int count, bool hide) => Editing.HideColumns(targetType, count, hide);
        public void GroupColumns(int count, bool group) => Editing.GroupColumns(count, group);
        public void AdjustRowHeight(int delta) => Editing.AdjustRowHeight(delta);
        public void AdjustColumnWidth(int delta) => Editing.AdjustColumnWidth(delta);
        public void MoveActiveCellBy(int rowDelta, int columnDelta) => Editing.MoveActiveCellBy(rowDelta, columnDelta);
        public void ActivateAdjacentSheet(int steps, bool forward) => Editing.ActivateAdjacentSheet(steps, forward);
        public void MoveActiveSheet(int steps, bool forward) => Editing.MoveActiveSheet(steps, forward);
        public void PasteEntireRows(Excel.Range source, int copies, bool pasteAfter) => Editing.PasteEntireRows(source, copies, pasteAfter);
        public void PasteEntireColumns(Excel.Range source, int copies, bool pasteAfter) => Editing.PasteEntireColumns(source, copies, pasteAfter);
        public void ResizeSelection(int up, int down, int left, int right) => Editing.ResizeSelection(up, down, left, right);
        public void ScrollHalf(bool scrollUp, int repeatCount) => Editing.ScrollHalf(scrollUp, repeatCount);
        public void ScrollHalfHorizontal(bool scrollLeft, int repeatCount) => Editing.ScrollHalfHorizontal(scrollLeft, repeatCount);
        public void EnsureActiveCellVisible() => Editing.EnsureActiveCellVisible();
        public void ScrollActiveRowToTop(double offsetPoints) => Editing.ScrollActiveRowToTop(offsetPoints);
        public void ScrollActiveRowToBottom(double offsetPoints) => Editing.ScrollActiveRowToBottom(offsetPoints);
        public void ScrollActiveRowToMiddle() => Editing.ScrollActiveRowToMiddle();
        public void ScrollActiveColumnToLeft() => Editing.ScrollActiveColumnToLeft();
        public void ScrollActiveColumnToRight() => Editing.ScrollActiveColumnToRight();
        public void ScrollActiveColumnToCenter() => Editing.ScrollActiveColumnToCenter();
        public void DeleteActiveCellComment() => Editing.DeleteActiveCellComment();
        public void DeleteAllComments() => Editing.DeleteAllComments();
        public void ToggleActiveCommentVisibility() => Editing.ToggleActiveCommentVisibility();
        public void ShowActiveComment() => Editing.ShowActiveComment();
        public void HideActiveComment() => Editing.HideActiveComment();
        public void ToggleAllCommentsVisibility() => Editing.ToggleAllCommentsVisibility();
        public void SetCommentIndicatorMode(int mode) => Editing.SetCommentIndicatorMode(mode);
        public void NavigateComments(bool forward, int steps) => Editing.NavigateComments(forward, steps);
        #endregion

        #region Workbook maintenance
        public void ClearUnnecessaryFormatting() => Optimizer.ClearUnnecessaryFormatting();
        public void DrawDependencyMap() => Analysis.DrawDependencyMap();
#endregion

        #region Trace dialogs
        public void TracePrecedentsDialog() => TraceDialogs.ShowPrecedentsDialog();
        public void TraceDependentsDialog() => TraceDialogs.ShowDependentsDialog();
        #endregion

        #region Chart helpers
        public void SelectNearestChart() => Charts.SelectNearestChart();
        public bool MoveSelectedLabels(double dx, double dy) => Charts.MoveSelectedLabels(dx, dy);
        public bool MoveSelectedChart(double dx, double dy) => Charts.MoveSelectedChart(dx, dy);
        #endregion
    }
}


