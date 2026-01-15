using System;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class AutoColorService
    {
        private readonly Excel.Application _app;
        private static readonly Regex SheetRefRegex = new Regex("('([^']+)'|[A-Za-z0-9_]+)!", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly int DefaultFontColor = ColorTranslator.ToOle(Color.Black);
        private static readonly int ExternalRefColor = ColorTranslator.ToOle(Color.FromArgb(120, 33, 112));
        private static readonly int OtherSheetColor = ColorTranslator.ToOle(Color.FromArgb(0, 128, 0));
        private static readonly int HardcodedNumericColor = ColorTranslator.ToOle(Color.FromArgb(0, 0, 255));

        public AutoColorService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public void ApplyAutoColor(Excel.Range target, int maxCells)
        {
            if (!RangeHelpers.IsRangeValid(target))
            {
                return;
            }

            long cellCount;
            try
            {
                cellCount = Convert.ToInt64(target.CountLarge);
            }
            catch
            {
                return;
            }

            if (cellCount <= 0)
            {
                return;
            }

            if (maxCells > 0 && cellCount > maxCells)
            {
                return;
            }

            Excel.Worksheet sheet = null;
            Excel.Workbook workbook = null;
            string sheetName = string.Empty;

            try
            {
                sheet = target.Worksheet;
                workbook = sheet?.Parent as Excel.Workbook;
                sheetName = sheet?.Name ?? string.Empty;

                if (sheet != null)
                {
                    if (target.Columns.Count >= sheet.Columns.Count || target.Rows.Count >= sheet.Rows.Count)
                    {
                        return;
                    }
                }
            }
            catch
            {
                sheetName = string.Empty;
            }

            using (new UiGuard(_app, hideStatusBar: true))
            {
                foreach (Excel.Range cell in target.Cells)
                {
                    try
                    {
                        ApplyColorToCell(cell, sheetName, workbook);
                    }
                    catch
                    {
                        // ignore cell errors
                    }
                    finally
                    {
                        ReleaseCom(cell);
                    }
                }
            }
        }

        private void ApplyColorToCell(Excel.Range cell, string sheetName, Excel.Workbook workbook)
        {
            if (!RangeHelpers.IsRangeValid(cell))
            {
                return;
            }

            object value;
            try
            {
                value = cell.Value2;
            }
            catch
            {
                return;
            }

            if (IsErrorValue(value))
            {
                return;
            }

            string cellText = Convert.ToString(value);
            if (string.IsNullOrWhiteSpace(cellText))
            {
                SetFontColorIfNeeded(cell, DefaultFontColor);
                return;
            }

            if (!IsNumericValue(value, cellText))
            {
                return;
            }

            bool hasFormula = false;
            try { hasFormula = Convert.ToBoolean(cell.HasFormula); } catch { hasFormula = false; }

            int desiredColor = DefaultFontColor;
            if (hasFormula)
            {
                string formulaText = null;
                try { formulaText = Convert.ToString(cell.Formula); } catch { formulaText = null; }

                if (IsExternalReference(formulaText))
                {
                    desiredColor = ExternalRefColor;
                }
                else if (IsOtherSheetReference(formulaText, sheetName, workbook))
                {
                    desiredColor = OtherSheetColor;
                }
                else
                {
                    desiredColor = DefaultFontColor;
                }
            }
            else
            {
                desiredColor = HardcodedNumericColor;
            }

            SetFontColorIfNeeded(cell, desiredColor);
        }

        private static bool IsNumericValue(object value, string text)
        {
            if (value == null)
            {
                return false;
            }

            if (value is double || value is float || value is decimal || value is int || value is long || value is short || value is byte)
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            double parsed;
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out parsed))
            {
                return true;
            }

            return double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out parsed);
        }

        private static bool IsErrorValue(object value)
        {
            if (value == null)
            {
                return false;
            }

            if (value is ErrorWrapper)
            {
                return true;
            }

            return false;
        }

        private static bool IsExternalReference(string formulaText)
        {
            if (string.IsNullOrEmpty(formulaText))
            {
                return false;
            }

            return formulaText.IndexOf('[') >= 0 && formulaText.IndexOf(']') >= 0;
        }

        private static bool IsOtherSheetReference(string formulaText, string currentSheet, Excel.Workbook workbook)
        {
            if (string.IsNullOrEmpty(formulaText))
            {
                return false;
            }

            if (formulaText.IndexOf('!') < 0)
            {
                return false;
            }

            if (workbook == null)
            {
                return false;
            }

            foreach (Match match in SheetRefRegex.Matches(formulaText))
            {
                string candidate = NormalizeSheetToken(match.Value);
                if (string.IsNullOrEmpty(candidate))
                {
                    continue;
                }

                if (candidate.IndexOf('[') >= 0 || candidate.IndexOf(']') >= 0)
                {
                    continue;
                }

                Excel.Worksheet ws = null;
                try
                {
                    ws = workbook.Worksheets[candidate] as Excel.Worksheet;
                }
                catch
                {
                    ws = null;
                }

                if (ws != null)
                {
                    try
                    {
                        if (!string.Equals(ws.Name, currentSheet, StringComparison.OrdinalIgnoreCase))
                        {
                            ReleaseCom(ws);
                            return true;
                        }
                    }
                    finally
                    {
                        ReleaseCom(ws);
                    }
                }
            }

            return false;
        }

        private static string NormalizeSheetToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return string.Empty;
            }

            token = token.Trim();
            if (token.EndsWith("!", StringComparison.Ordinal))
            {
                token = token.Substring(0, token.Length - 1);
            }

            token = token.Trim();
            if (token.Length >= 2 && token.StartsWith("'", StringComparison.Ordinal) && token.EndsWith("'", StringComparison.Ordinal))
            {
                token = token.Substring(1, token.Length - 2).Replace("''", "'");
            }

            return token.Trim();
        }

        private static void SetFontColorIfNeeded(Excel.Range cell, int desiredColor)
        {
            try
            {
                int currentColor = Convert.ToInt32(cell.Font.Color);
                if (currentColor != desiredColor)
                {
                    cell.Font.Color = desiredColor;
                }
            }
            catch
            {
                // ignore font errors
            }
        }

        private static void ReleaseCom(object comObject)
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
    }
}
