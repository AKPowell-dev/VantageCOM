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

            try
            {
                sheet.Activate();
            }
            catch
            {
                // ignored
            }
        }
    }
}
