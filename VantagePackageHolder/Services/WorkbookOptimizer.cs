using System;
using System.Diagnostics;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class WorkbookOptimizer
    {
        private readonly Excel.Application _app;

        public WorkbookOptimizer(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public void ClearUnnecessaryFormatting()
        {
            var workbook = _app.ActiveWorkbook;
            if (workbook == null)
            {
                return;
            }

            using (new UiGuard(_app, hideStatusBar: true))
            {
                var originalCalc = _app.Calculation;
                Excel.Worksheet originalSheet = null;
                string originalAddress = string.Empty;

                try
                {
                    originalSheet = _app.ActiveSheet as Excel.Worksheet;
                    if (_app.Selection is Excel.Range sel)
                    {
                        var firstCell = sel.Cells[1, 1] as Excel.Range;
                        if (firstCell != null)
                        {
                            originalAddress = firstCell.Address[false, false];
                        }
                    }
                }
                catch
                {
                    // ignore
                }

                _app.Calculation = Excel.XlCalculation.xlCalculationManual;
                _app.StatusBar = "Optimizing workbook...";
                _app.CutCopyMode = Excel.XlCutCopyMode.xlCopy;

                CleanupBrokenNames(workbook);

                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.ProtectContents)
                    {
                        continue;
                    }

                    int lastRow = GetLastUsedRow(sheet);
                    int lastCol = GetLastUsedColumn(sheet);

                    sheet.DisplayPageBreaks = false;

                    if (lastRow <= 0 || lastCol <= 0)
                    {
                        sheet.Cells.ClearFormats();
                        sheet.Cells.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                        continue;
                    }

                    if (lastRow < sheet.Rows.Count)
                    {
                        var rows = sheet.Range[sheet.Rows[lastRow + 1], sheet.Rows[sheet.Rows.Count]];
                        rows.ClearFormats();
                        rows.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                        rows.FormatConditions.Delete();
                    }

                    if (lastCol < sheet.Columns.Count)
                    {
                        var cols = sheet.Range[sheet.Columns[lastCol + 1], sheet.Columns[sheet.Columns.Count]];
                        cols.ClearFormats();
                        cols.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                        cols.FormatConditions.Delete();
                    }

                    TrimConditionalFormatting(sheet);

                    try
                    {
                        var used = sheet.UsedRange;
                        _ = used.Value;
                    }
                    catch
                    {
                        // ignore
                    }
                }

                if (workbook.Connections.Count > 0)
                {
                    try
                    {
                        workbook.RefreshAll();
                    }
                    catch
                    {
                        // ignore connection refresh errors
                    }
                }

                RefreshExcelCaches(workbook);

                if (originalSheet != null)
                {
                    RangeHelpers.SafeActivateSheet(originalSheet);
                    if (!string.IsNullOrEmpty(originalAddress))
                    {
                        try
                        {
                            var rng = originalSheet.Range[originalAddress];
                            RangeHelpers.SafeSelect(rng);
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }

                _app.Calculation = originalCalc;
                _app.StatusBar = false;

                if (ShouldRestartToFlushCaches(workbook))
                {
                    return;
                }

                RunMacroIfExists("SetStatusBarTemporarily", "Workbook formatting cache cleared.", 2500);
            }
        }

        private void CleanupBrokenNames(Excel.Workbook workbook)
        {
            foreach (Excel.Name nm in workbook.Names)
            {
                try
                {
                    var _ = nm.RefersTo;
                }
                catch
                {
                    try
                    {
                        nm.Delete();
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                foreach (Excel.Name name in sheet.Names)
                {
                    try
                    {
                        var _ = name.RefersTo;
                    }
                    catch
                    {
                        try
                        {
                            name.Delete();
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
            }
        }

        private void TrimConditionalFormatting(Excel.Worksheet sheet)
        {
            try
            {
                if (sheet.Cells.FormatConditions.Count > 0)
                {
                    sheet.Cells.FormatConditions.Delete();
                }
            }
            catch
            {
                // ignore
            }
        }

        private int GetLastUsedRow(Excel.Worksheet sheet)
        {
            try
            {
                var lastCell = sheet.Cells.Find(
                    What: "*",
                    After: sheet.Cells[1, 1],
                    LookIn: Excel.XlFindLookIn.xlFormulas,
                    LookAt: Excel.XlLookAt.xlPart,
                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Excel.XlSearchDirection.xlPrevious,
                    MatchCase: false);
                if (lastCell != null)
                {
                    return lastCell.Row;
                }
            }
            catch
            {
                // ignore
            }

            return 0;
        }

        private int GetLastUsedColumn(Excel.Worksheet sheet)
        {
            try
            {
                var lastCell = sheet.Cells.Find(
                    What: "*",
                    After: sheet.Cells[1, 1],
                    LookIn: Excel.XlFindLookIn.xlFormulas,
                    LookAt: Excel.XlLookAt.xlPart,
                    SearchOrder: Excel.XlSearchOrder.xlByColumns,
                    SearchDirection: Excel.XlSearchDirection.xlPrevious,
                    MatchCase: false);
                if (lastCell != null)
                {
                    return lastCell.Column;
                }
            }
            catch
            {
                // ignore
            }

            return 0;
        }

        private void RefreshExcelCaches(Excel.Workbook workbook)
        {
            using (new UiGuard(_app, hideStatusBar: true))
            {
                _app.StatusBar = "Refreshing workbook caches...";
                _app.CutCopyMode = Excel.XlCutCopyMode.xlCopy;

                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    ws.DisplayPageBreaks = false;
                }

                foreach (Excel.PivotCache cache in workbook.PivotCaches())
                {
                    try
                    {
                        cache.MissingItemsLimit = Excel.XlPivotTableMissingItems.xlMissingItemsNone;
                        cache.Refresh();
                    }
                    catch
                    {
                        // ignore
                    }
                }

                workbook.Application.CalculateFullRebuild();
                _app.StatusBar = false;
            }
        }

        private bool ShouldRestartToFlushCaches(Excel.Workbook workbook)
        {
            var result = MessageBox.Show(
                "Optimization complete." + Environment.NewLine +
                "Reopen Excel with this workbook to flush caches?" + Environment.NewLine + Environment.NewLine +
                "This will save the workbook and close the current Excel session.",
                "Reopen workbook",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            return result == DialogResult.Yes && RestartWorkbookFresh(workbook);
        }

        private bool RestartWorkbookFresh(Excel.Workbook workbook)
        {
            try
            {
                var path = workbook.FullName;
                if (string.IsNullOrWhiteSpace(path))
                {
                    MessageBox.Show("Please save the workbook before running the refresh.", "Reopen workbook", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }

                var excelPath = System.IO.Path.Combine(_app.Path, "EXCEL.EXE");
                _app.StatusBar = "Reopening workbook...";
                workbook.Save();
                Process.Start(new ProcessStartInfo(excelPath, "\"" + path + "\"")
                {
                    UseShellExecute = true
                });
                _app.Quit();
                return true;
            }
            catch
            {
                MessageBox.Show("Could not restart Excel automatically. Please reopen the workbook manually.", "Reopen workbook", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                _app.StatusBar = false;
                return false;
            }
        }

        private void RunMacroIfExists(string name, params object[] args)
        {
            try
            {
                if (args == null || args.Length == 0)
                {
                    _app.Run(name);
                }
                else
                {
                    _app.Run(name, args);
                }
            }
            catch
            {
                // ignore
            }
        }
    }
}
