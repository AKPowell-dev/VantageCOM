using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    [ComVisible(true)]
    [Guid("7D759476-0E72-4B44-B296-FFACDC61CCAA")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class VantageEngine : IDisposable
    {
        private readonly Excel.Application _excel;
        private readonly ClipboardService _clipboard;
        private readonly FormatService _format;
        private readonly WorkbookOptimizer _optimizer;
        private readonly EditingService _editing;
        private readonly WorkbookAnalysisService _analysis;
        private readonly ChartNavigator _charts;

        public VantageEngine(Excel.Application excel)
        {
            _excel = excel ?? throw new ArgumentNullException(nameof(excel));
            _clipboard = new ClipboardService(excel);
            var pptExporter = new PowerPointExporter();
            _format = new FormatService(excel, _clipboard, pptExporter);
            _optimizer = new WorkbookOptimizer(excel);
            _editing = new EditingService(excel);
            _analysis = new WorkbookAnalysisService(excel);
            _charts = new ChartNavigator(excel);
        }

        public void Dispose() { }

        #region Clipboard hooks
        public void ClipboardHandleCopy() => _clipboard.HandleCopy();
        public void ClipboardHandleCut() => _clipboard.HandleCut();
        public void ClipboardHandlePaste() => _clipboard.HandlePaste();
        public void ClipboardHandlePasteValues() => _clipboard.HandlePasteValues();
        public void ClipboardHandlePasteFormulas() => _clipboard.HandlePasteFormulas();
        public void ClipboardOpenPasteSpecial() => _clipboard.OpenPasteSpecialDialog();
        public Excel.Range ClipboardGetCopyRange() => _clipboard.GetCopyRange();
        public void ClipboardSetCopyRange(Excel.Range range) => _clipboard.SetCopyRange(range);
        #endregion

        #region Formatting helpers
        public void PasteExact() => _format.PasteExact();
        public void PasteCondensed() => _format.PasteCondensed();
        public void SmartFillRight() => _format.SmartFillRight();
        public void SmartFormatRight() => _format.SmartFormatRight();
        public void OutlineSelectionHighlight() => _format.OutlineSelectionHighlight();
        public void SmartFillDown() => _format.SmartFillDown();
        public void WrapFormulaWithCircCheck() => _format.WrapFormulaWithCircCheck();
        public bool CopySelectionAsPicturePrintSafe() => _format.CopySelectionAsPicturePrintSafe();
        public void CopySelectionToPowerPoint() => _format.CopyPasteSelectionToPowerPoint();
        public void FormatChartFg() => _format.FormatChartFg();
        public void ApplyInteriorColor(bool isNull, bool isTheme, int themeColor, double tint, int rgb) => _format.ApplyInteriorColor(isNull, isTheme, themeColor, tint, rgb);
        public bool ApplyFontColor(bool isNull, bool isTheme, int themeColor, int objectThemeColor, double tint, int rgb) => _format.ApplyFontColor(isNull, isTheme, themeColor, objectThemeColor, tint, rgb);
        public void ApplyShapeFillColor(bool isNull, bool isTheme, int themeColor, double tint, int rgb) => _format.ApplyShapeFillColor(isNull, isTheme, themeColor, tint, rgb);
        public void ApplyShapeFontColor(bool isNull, bool isTheme, int themeColor, int objectThemeColor, double tint, int rgb) => _format.ApplyShapeFontColor(isNull, isTheme, themeColor, objectThemeColor, tint, rgb);
        public void ApplyShapeBorderColor(bool isNull, bool isTheme, int themeColor, double tint, int rgb) => _format.ApplyShapeBorderColor(isNull, isTheme, themeColor, tint, rgb);
        public bool ApplySmartFontColor(bool isNull, bool isTheme, int themeColor, int objectThemeColor, double tint, int rgb) => _format.ApplySmartFontColor(isNull, isTheme, themeColor, objectThemeColor, tint, rgb);
        public void ApplySmartFillColor(bool isNull, bool isTheme, int themeColor, double tint, int rgb) => _format.ApplySmartFillColor(isNull, isTheme, themeColor, tint, rgb);

        public void LockCellReference() => _format.LockCellReference();
        public void ResetCycleState() => _format.ResetCycleState();
        public void ClearFormatting() => _format.ClearFormatting();
        public void CycleFormatting() => _format.CycleFormatting();
        public void CycleNumberFormat() => _format.CycleNumberFormat();
        public void BinaryCycle() => _format.BinaryCycle();
        public void YearDisplayCycle() => _format.YearDisplayCycle();
        public void NumberNarrativeCycle() => _format.NumberNarrativeCycle();
        public void PercentCycle() => _format.PercentCycle();
        public void FlipSign() => _format.FlipSign();
        public void ReverseSelectionOrder() => _format.ReverseSelectionOrder();
        public void TrimConditionalFormatting() => _format.TrimConditionalFormatting();
        public void CurrencyCycle() => _format.CurrencyCycle();
        public void ToggleBorder(string targetKey, int lineStyle, int weight) => _format.ToggleBorder(targetKey, lineStyle, weight);
        public void DeleteBorder(string targetKey) => _format.DeleteBorder(targetKey);
        public void SetBorderColor(string targetKey, bool isNull, bool isTheme, int themeColor, double tintAndShade, int rgb) => _format.SetBorderColor(targetKey, isNull, isTheme, themeColor, tintAndShade, rgb);

        public void AdjustNumbers(int count, bool subtract, bool grow) => _editing.AdjustNumbers(count, subtract, grow);
        public void ApplyAutoFill() => _editing.ApplyAutoFill();
        public void NavigateSpecialCells(int typeValue, Excel.XlSearchOrder searchOrder, bool forward, int steps) => _editing.NavigateSpecialCells(typeValue, searchOrder, forward, steps);
        public void SubstituteType(string text) => _editing.SubstituteType(text);
        public void InsertRows(int count, bool append) => _editing.InsertRows(count, append);
        public void DeleteRows(int targetType, int count) => _editing.DeleteRows(targetType, count);
        public void HideRows(int targetType, int count, bool hide) => _editing.HideRows(targetType, count, hide);
        public void GroupRows(int count, bool group) => _editing.GroupRows(count, group);
        public void InsertColumns(int count, bool append) => _editing.InsertColumns(count, append);
        public void DeleteColumns(int targetType, int count) => _editing.DeleteColumns(targetType, count);
        public void HideColumns(int targetType, int count, bool hide) => _editing.HideColumns(targetType, count, hide);
        public void GroupColumns(int count, bool group) => _editing.GroupColumns(count, group);
        #endregion

        #region Workbook maintenance
        public void ClearUnnecessaryFormatting() => _optimizer.ClearUnnecessaryFormatting();
        public void DrawDependencyMap() => _analysis.DrawDependencyMap();
        #endregion

        #region Chart helpers
        public void SelectNearestChart() => _charts.SelectNearestChart();
        public bool MoveSelectedLabels(double dx, double dy) => _charts.MoveSelectedLabels(dx, dy);
        public bool MoveSelectedChart(double dx, double dy) => _charts.MoveSelectedChart(dx, dy);
        #endregion
    }
}


