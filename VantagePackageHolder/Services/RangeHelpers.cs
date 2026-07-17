using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal static class RangeHelpers
    {
        public static bool TryGetActiveRange(Excel.Application app, out Excel.Range range)
        {
            range = null;
            if (app == null)
            {
                return false;
            }

            try
            {
                if (app.Selection is Excel.Range selected)
                {
                    range = selected;
                    return true;
                }
            }
            catch
            {
                range = null;
            }

            return false;
        }

        public static bool TryGetRangeOrActiveCell(Excel.Application app, out Excel.Range range)
        {
            if (TryGetActiveRange(app, out range))
            {
                return true;
            }

            range = null;
            if (app == null)
            {
                return false;
            }

            try
            {
                if (app.ActiveCell is Excel.Range cell)
                {
                    range = cell;
                    return true;
                }
            }
            catch
            {
                range = null;
            }

            return false;
        }

        public static bool IsRangeValid(Excel.Range range)
        {
            if (range == null)
            {
                return false;
            }

            try
            {
                _ = range.Address[false, false];
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static string BuildRangeKey(Excel.Range range)
        {
            if (!IsRangeValid(range))
            {
                return string.Empty;
            }

            try
            {
                var workbook = range.Worksheet?.Parent as Excel.Workbook;
                var wbName = workbook?.FullName ?? workbook?.Name ?? string.Empty;
                var sheetName = range.Worksheet?.Name ?? string.Empty;
                var address = range.Address[false, false, Excel.XlReferenceStyle.xlA1];
                return $"{wbName}|{sheetName}|{address}";
            }
            catch
            {
                return string.Empty;
            }
        }

        public static void SafeSelect(Excel.Range range)
        {
            if (!IsRangeValid(range))
            {
                return;
            }

            try
            {
                range.Select();
            }
            catch
            {
                // ignored
            }
        }

        public static void SafeActivateSheet(Excel.Worksheet sheet)
        {
            if (sheet == null)
            {
                return;
            }

            // If this sheet is already the active sheet of the active window, do
            // not re-activate it. With multiple windows of the same workbook,
            // Worksheet.Activate hops focus to whichever sibling window already
            // shows the sheet — so re-activating the sheet you're already on jumps
            // you to another window. Selecting works fine without re-activating.
            try
            {
                var win = sheet.Application?.ActiveWindow;
                if (win != null && win.ActiveSheet is Excel.Worksheet current && SameWorksheet(current, sheet))
                {
                    return;
                }
            }
            catch
            {
                // fall through and activate
            }

            try
            {
                sheet.Activate();
            }
            catch
            {
                // ignored
            }
        }

        private static bool SameWorksheet(Excel.Worksheet a, Excel.Worksheet b)
        {
            if (a == null || b == null)
            {
                return false;
            }

            try
            {
                if (!string.Equals(a.Name, b.Name, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                var wbA = (a.Parent as Excel.Workbook)?.Name;
                var wbB = (b.Parent as Excel.Workbook)?.Name;
                return string.Equals(wbA, wbB, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }
    }
}
