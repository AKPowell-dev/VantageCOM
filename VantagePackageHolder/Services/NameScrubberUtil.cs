using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal static class NameScrubberUtil
    {
        private static readonly string[] ErrorTokens =
        {
            "#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!"
        };

        public static bool IsNative(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            return name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("_xlda.", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsNative(Excel.Name name)
        {
            if (name == null)
            {
                return false;
            }

            return IsNative(name.Name);
        }

        public static bool IsLinked(Excel.Name name)
        {
            if (name == null)
            {
                return false;
            }

            var refersTo = SafeRefersTo(name);
            return refersTo.IndexOf("[", StringComparison.OrdinalIgnoreCase) >= 0
                && refersTo.IndexOf("]", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool IsErroneous(Excel.Name name)
        {
            if (name == null)
            {
                return false;
            }

            var refersTo = SafeRefersTo(name);
            if (string.IsNullOrEmpty(refersTo))
            {
                return false;
            }

            foreach (var token in ErrorTokens)
            {
                if (refersTo.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        public static bool IsLambda(Excel.Name name)
        {
            if (name == null)
            {
                return false;
            }

            var refersTo = SafeRefersTo(name);
            return refersTo.StartsWith("=LAMBDA", StringComparison.OrdinalIgnoreCase);
        }

        public static string SafeRefersTo(Excel.Name name)
        {
            if (name == null)
            {
                return string.Empty;
            }

            try
            {
                return name.RefersTo?.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public static Excel.Range TryGetRefersToRange(Excel.Name name)
        {
            if (name == null)
            {
                return null;
            }

            try
            {
                return name.RefersToRange;
            }
            catch
            {
                return null;
            }
        }

        public static bool FormulaReferencesName(Excel.Range range, Excel.Name name)
        {
            if (range == null || name == null)
            {
                return false;
            }

            string formula = string.Empty;
            try
            {
                formula = range.Formula?.ToString() ?? string.Empty;
            }
            catch
            {
                return false;
            }

            if (string.IsNullOrEmpty(formula))
            {
                return false;
            }

            var nameToken = name.Name;
            if (string.IsNullOrEmpty(nameToken))
            {
                return false;
            }

            var shortName = StripSheetPrefix(nameToken);
            return RegexMatchesToken(formula, nameToken) || (shortName != nameToken && RegexMatchesToken(formula, shortName));
        }

        public static bool HasDependents(Excel.Name name, Func<bool> shouldCancel)
        {
            if (name == null)
            {
                return false;
            }

            if (IsNative(name) || IsLinked(name))
            {
                return true;
            }

            var dependents = GetDependentsAcrossWorkbook(name, shouldCancel);
            return dependents.Count > 0;
        }

        public static List<Excel.Range> GetDependents(Excel.Range refersToRange, Excel.Name name, Func<bool> shouldCancel)
        {
            var results = new List<Excel.Range>();
            if (refersToRange == null)
            {
                return results;
            }

            Excel.Range anchor = null;
            Excel.Workbook wb = null;
            Excel.Application app = null;

            try
            {
                app = refersToRange.Application;
                wb = refersToRange.Worksheet?.Parent as Excel.Workbook;
                if (wb == null)
                {
                    return results;
                }

                if (Convert.ToBoolean(refersToRange.MergeCells))
                {
                    return results;
                }

                anchor = refersToRange.Cells[1, 1] as Excel.Range;
                if (anchor == null)
                {
                    return results;
                }

                var displayObjects = wb.DisplayDrawingObjects;
                wb.DisplayDrawingObjects = Excel.XlDisplayDrawingObjects.xlHide;

                try
                {
                    anchor.ShowDependents(Type.Missing);
                }
                catch
                {
                    // ignore
                }

                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var stopwatch = Stopwatch.StartNew();

                for (int arrow = 1; arrow < 256; arrow++)
                {
                    if (shouldCancel != null && shouldCancel())
                    {
                        break;
                    }
                    if (stopwatch.Elapsed.TotalSeconds > 8)
                    {
                        break;
                    }

                    Excel.Range dep = null;
                    try
                    {
                        dep = anchor.NavigateArrow(false, arrow, 1);
                    }
                    catch
                    {
                        dep = null;
                    }

                    if (dep == null || IsSameAddress(dep, anchor))
                    {
                        break;
                    }

                    AddDependent(results, seen, dep);

                    for (int link = 2; link < 256; link++)
                    {
                        if (shouldCancel != null && shouldCancel())
                        {
                            break;
                        }
                        if (stopwatch.Elapsed.TotalSeconds > 8)
                        {
                            break;
                        }

                        Excel.Range dep2 = null;
                        try
                        {
                            dep2 = anchor.NavigateArrow(false, arrow, link);
                        }
                        catch
                        {
                            dep2 = null;
                        }

                        if (dep2 == null || IsSameAddress(dep2, anchor))
                        {
                            break;
                        }

                        AddDependent(results, seen, dep2);
                    }
                }

                try
                {
                    anchor.ShowDependents(true);
                }
                catch
                {
                    // ignore
                }

                try
                {
                    wb.DisplayDrawingObjects = displayObjects;
                }
                catch
                {
                    // ignore
                }

                TryClearArrows(app);
            }
            catch
            {
                // ignore
            }

            return results;
        }

        public static List<Excel.Range> GetDependentsAcrossWorkbook(Excel.Name name, Func<bool> shouldCancel)
        {
            var results = new List<Excel.Range>();
            if (name == null)
            {
                return results;
            }

            Excel.Workbook wb = null;
            try
            {
                if (name.Parent is Excel.Workbook namedWorkbook)
                {
                    wb = namedWorkbook;
                }
                else if (name.Parent is Excel.Worksheet namedSheet)
                {
                    wb = namedSheet.Parent as Excel.Workbook;
                }
            }
            catch
            {
                wb = null;
            }

            if (wb == null)
            {
                return results;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (shouldCancel != null && shouldCancel())
                {
                    break;
                }

                Excel.Range formulaCells = null;
                try
                {
                    formulaCells = ws.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                }
                catch
                {
                    formulaCells = null;
                }

                if (formulaCells == null)
                {
                    continue;
                }

                foreach (Excel.Range area in formulaCells.Areas)
                {
                    foreach (Excel.Range cell in area.Cells)
                    {
                        if (shouldCancel != null && shouldCancel())
                        {
                            return results;
                        }

                        if (!FormulaReferencesName(cell, name))
                        {
                            continue;
                        }

                        AddDependent(results, seen, cell);
                    }
                }
            }

            return results;
        }

        public static List<Excel.Range> ApplyNameToDependents(Excel.Name name, Func<bool> shouldCancel, Action<int, int> reportProgress)
        {
            var changed = new List<Excel.Range>();
            if (name == null)
            {
                return changed;
            }

            var deps = GetDependentsAcrossWorkbook(name, shouldCancel);
            int count = deps.Count;
            int index = 0;

            foreach (var dep in deps)
            {
                index++;
                if (shouldCancel != null && shouldCancel())
                {
                    break;
                }
                reportProgress?.Invoke(index, count);

                if (dep == null)
                {
                    continue;
                }

                string originalFormula;
                object originalValue;
                bool hasArray = false;
                try
                {
                    hasArray = Convert.ToBoolean(dep.HasArray);
                }
                catch
                {
                    hasArray = false;
                }

                try
                {
                    originalFormula = hasArray ? dep.FormulaArray?.ToString() ?? string.Empty : dep.Formula?.ToString() ?? string.Empty;
                }
                catch
                {
                    originalFormula = string.Empty;
                }

                try
                {
                    originalValue = dep.Value2;
                }
                catch
                {
                    originalValue = null;
                }

                try
                {
                    dep.ApplyNames(new[] { name.Name }, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlApplyNamesOrder.xlRowThenColumn, Type.Missing);
                }
                catch
                {
                    continue;
                }

                try
                {
                    if (!ValuesEqual(dep.Value2, originalValue))
                    {
                        if (hasArray)
                        {
                            dep.FormulaArray = originalFormula;
                        }
                        else
                        {
                            dep.Formula = originalFormula;
                        }
                    }
                }
                catch
                {
                    // ignore
                }

                try
                {
                    var newFormula = hasArray ? dep.FormulaArray?.ToString() ?? string.Empty : dep.Formula?.ToString() ?? string.Empty;
                    if (!string.Equals(newFormula, originalFormula, StringComparison.OrdinalIgnoreCase))
                    {
                        changed.Add(dep);
                    }
                }
                catch
                {
                    // ignore
                }
            }

            return changed;
        }

        public static List<Excel.Range> UnapplyNameFromDependents(Excel.Name name, Func<bool> shouldCancel, Action<int, int> reportProgress)
        {
            var changed = new List<Excel.Range>();
            if (name == null)
            {
                return changed;
            }

            var refersToRangeText = SafeRefersTo(name);
            if (string.IsNullOrWhiteSpace(refersToRangeText))
            {
                return changed;
            }

            var deps = GetDependentsAcrossWorkbook(name, shouldCancel);
            int count = deps.Count;
            int index = 0;

            foreach (var dep in deps)
            {
                index++;
                if (shouldCancel != null && shouldCancel())
                {
                    break;
                }
                reportProgress?.Invoke(index, count);

                if (dep == null)
                {
                    continue;
                }

                bool hasArray = false;
                try
                {
                    hasArray = Convert.ToBoolean(dep.HasArray);
                }
                catch
                {
                    hasArray = false;
                }

                string formula;
                try
                {
                    formula = hasArray ? dep.FormulaArray?.ToString() ?? string.Empty : dep.Formula?.ToString() ?? string.Empty;
                }
                catch
                {
                    continue;
                }

                if (string.IsNullOrEmpty(formula))
                {
                    continue;
                }

                string refersTo = refersToRangeText;
                if (refersTo.StartsWith("=", StringComparison.Ordinal))
                {
                    refersTo = refersTo.Substring(1);
                }

                refersTo = NormalizeRefersTo(dep.Application, refersTo);

                if (dep.Worksheet != null && name.Parent is Excel.Worksheet nameSheet &&
                    string.Equals(dep.Worksheet.Name, nameSheet.Name, StringComparison.OrdinalIgnoreCase))
                {
                    refersTo = StripSheetPrefix(refersTo);
                }

                var nameToken = name.Name;
                var shortName = StripSheetPrefix(nameToken);
                var updated = ReplaceNameToken(formula, nameToken, refersTo);
                if (shortName != nameToken)
                {
                    updated = ReplaceNameToken(updated, shortName, refersTo);
                }

                if (string.Equals(updated, formula, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                object originalValue = null;
                try
                {
                    originalValue = dep.Value2;
                }
                catch
                {
                    originalValue = null;
                }

                try
                {
                    if (hasArray)
                    {
                        dep.FormulaArray = updated;
                    }
                    else
                    {
                        dep.Formula = updated;
                    }
                }
                catch
                {
                    continue;
                }

                try
                {
                    if (!ValuesEqual(dep.Value2, originalValue))
                    {
                        if (hasArray)
                        {
                            dep.FormulaArray = formula;
                        }
                        else
                        {
                            dep.Formula = formula;
                        }
                        continue;
                    }
                }
                catch
                {
                    // ignore
                }

                try
                {
                    var newFormula = hasArray ? dep.FormulaArray?.ToString() ?? string.Empty : dep.Formula?.ToString() ?? string.Empty;
                    if (!string.Equals(newFormula, formula, StringComparison.OrdinalIgnoreCase))
                    {
                        changed.Add(dep);
                    }
                }
                catch
                {
                    // ignore
                }
            }

            return changed;
        }

        public static bool ExternalLinkMissing(string refersTo, string workbookPath)
        {
            if (string.IsNullOrEmpty(refersTo))
            {
                return false;
            }

            int start = refersTo.IndexOf("[", StringComparison.OrdinalIgnoreCase);
            int end = refersTo.IndexOf("]", StringComparison.OrdinalIgnoreCase);
            if (start < 0 || end <= start)
            {
                return false;
            }

            var filePart = refersTo.Substring(start + 1, end - start - 1).Trim();
            if (string.IsNullOrEmpty(filePart))
            {
                return false;
            }

            if (Uri.TryCreate(filePart, UriKind.Absolute, out var uri))
            {
                if (uri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) ||
                    uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase) ||
                    uri.Scheme.Equals("file", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            try
            {
                if (Path.IsPathRooted(filePart))
                {
                    return !File.Exists(filePart);
                }

                if (!string.IsNullOrEmpty(workbookPath))
                {
                    var full = Path.Combine(workbookPath, filePart);
                    return !File.Exists(full);
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        public static string StripSheetPrefix(string token)
        {
            if (string.IsNullOrEmpty(token))
            {
                return token;
            }

            int bang = token.LastIndexOf('!');
            if (bang >= 0 && bang + 1 < token.Length)
            {
                return token.Substring(bang + 1);
            }

            return token;
        }

        private static string NormalizeRefersTo(Excel.Application app, string refersTo)
        {
            if (string.IsNullOrEmpty(refersTo))
            {
                return refersTo;
            }

            try
            {
                return app.ConvertFormula(refersTo, Excel.XlReferenceStyle.xlA1, Excel.XlReferenceStyle.xlA1, Excel.XlReferenceType.xlAbsolute);
            }
            catch
            {
                return refersTo;
            }
        }

        private static string ReplaceNameToken(string formula, string nameToken, string replacement)
        {
            if (string.IsNullOrEmpty(formula) || string.IsNullOrEmpty(nameToken))
            {
                return formula;
            }

            var pattern = BuildTokenRegex(nameToken);
            return Regex.Replace(formula, pattern, replacement ?? string.Empty, RegexOptions.IgnoreCase);
        }

        private static bool RegexMatchesToken(string formula, string token)
        {
            if (string.IsNullOrEmpty(formula) || string.IsNullOrEmpty(token))
            {
                return false;
            }

            var pattern = BuildTokenRegex(token);
            return Regex.IsMatch(formula, pattern, RegexOptions.IgnoreCase);
        }

        private static string BuildTokenRegex(string token)
        {
            var escaped = Regex.Escape(token);
            return $@"(?<![A-Za-z0-9_\.]){escaped}(?![A-Za-z0-9_\.])";
        }

        private static void AddDependent(List<Excel.Range> results, HashSet<string> seen, Excel.Range dep)
        {
            if (dep == null)
            {
                return;
            }

            var key = RangeHelpers.BuildRangeKey(dep);
            if (string.IsNullOrEmpty(key))
            {
                return;
            }

            if (!seen.Add(key))
            {
                return;
            }

            results.Add(dep);
        }

        private static bool IsSameAddress(Excel.Range left, Excel.Range right)
        {
            if (left == null || right == null)
            {
                return false;
            }

            try
            {
                return string.Equals(left.Address[true, true, Excel.XlReferenceStyle.xlA1, true], right.Address[true, true, Excel.XlReferenceStyle.xlA1, true], StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static void TryClearArrows(Excel.Application app)
        {
            if (app == null)
            {
                return;
            }

            try
            {
                if (app.ActiveSheet is Excel.Worksheet sheet)
                {
                    sheet.ClearArrows();
                }
            }
            catch
            {
                // ignore
            }
        }

        private static bool ValuesEqual(object left, object right)
        {
            if (left == null && right == null)
            {
                return true;
            }
            if (left == null || right == null)
            {
                return false;
            }

            return left.Equals(right);
        }
    }
}
