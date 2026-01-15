using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class UtilService
    {
        private readonly Excel.Application _app;
        private readonly Stopwatch _stopwatch = new Stopwatch();
        private readonly Timer _statusTimer;
        private string _savedStatusMessage = string.Empty;
        private bool _tempVisible;
        private long _lastStatusTicks;

        private const int ProgressBarLength = 13;

        public UtilService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _stopwatch.Start();
            _statusTimer = new Timer();
            _statusTimer.Tick += OnStatusTimerTick;
        }

        public void TimeClear()
        {
            _stopwatch.Restart();
        }

        public double GetQueryPerformanceTime(string format)
        {
            if (string.IsNullOrEmpty(format))
            {
                format = "0.0000";
            }

            double elapsed = _stopwatch.Elapsed.TotalSeconds;
            int decimals = GetDecimalDigits(format);
            if (decimals >= 0)
            {
                elapsed = Math.Round(elapsed, decimals);
            }

            TimeClear();
            return elapsed;
        }

        public void SetStatusBar(string text, long currentCount, long maximumCount, double percent, int numDigitsAfterDecimal, bool progressBar, bool countPerMax)
        {
            try
            {
                if (text == null)
                {
                    text = string.Empty;
                }

                if (_tempVisible)
                {
                    _savedStatusMessage = text;
                    return;
                }

                if (string.IsNullOrEmpty(text))
                {
                    _savedStatusMessage = string.Empty;
                    _app.StatusBar = false;
                    return;
                }

                string progressText = string.Empty;
                if (progressBar)
                {
                    if (currentCount >= 0 && maximumCount >= currentCount && maximumCount > 0)
                    {
                        percent = (double)currentCount / maximumCount;
                    }

                    percent *= 100.0;
                    if (percent < 0 || percent > 100)
                    {
                        _app.StatusBar = false;
                        return;
                    }

                    progressText = BuildProgressText(
                        percent,
                        numDigitsAfterDecimal,
                        countPerMax ? currentCount : -1,
                        countPerMax ? maximumCount : -1);
                }

                if (!progressBar || ShouldUpdateStatus())
                {
                    _savedStatusMessage = text + progressText;
                    _app.StatusBar = _savedStatusMessage;
                    _lastStatusTicks = Stopwatch.GetTimestamp();
                }
            }
            catch
            {
            }
        }

        public void SetStatusBarTemporarily(string text, int milliseconds, bool disablePrefix, string statusPrefix)
        {
            try
            {
            if (text == null)
            {
                text = string.Empty;
            }

                _tempVisible = true;
                if (_statusTimer.Enabled)
                {
                    _statusTimer.Stop();
                }

                _statusTimer.Interval = Math.Max(1, milliseconds);
                _statusTimer.Start();

                if (!disablePrefix && !string.IsNullOrEmpty(statusPrefix))
                {
                    _app.StatusBar = statusPrefix + text;
                }
                else
                {
                    _app.StatusBar = text;
                }
            }
            catch
            {
            }
        }

        public bool RegExpMatch(string str, string matchPattern, bool isIgnoreCase, bool isGlobal, bool isMultiline)
        {
            if (str == null)
            {
                str = string.Empty;
            }

            if (matchPattern == null)
            {
                matchPattern = string.Empty;
            }

            var options = RegexOptions.CultureInvariant;
            if (isIgnoreCase)
            {
                options |= RegexOptions.IgnoreCase;
            }

            if (isMultiline)
            {
                options |= RegexOptions.Multiline;
            }

            return Regex.IsMatch(str, matchPattern, options);
        }

        public string RegExpSearch(string str, string matchPattern, bool isIgnoreCase, bool isGlobal, bool isMultiline)
        {
            if (str == null)
            {
                str = string.Empty;
            }

            if (matchPattern == null)
            {
                matchPattern = string.Empty;
            }

            var options = RegexOptions.CultureInvariant;
            if (isIgnoreCase)
            {
                options |= RegexOptions.IgnoreCase;
            }

            if (isMultiline)
            {
                options |= RegexOptions.Multiline;
            }

            var match = Regex.Match(str, matchPattern, options);
            return match.Success ? match.Value : string.Empty;
        }

        public string RegExpReplace(string str, string matchPattern, string replaceStr, bool isIgnoreCase, bool isGlobal, bool isMultiline)
        {
            if (str == null)
            {
                str = string.Empty;
            }

            if (matchPattern == null)
            {
                matchPattern = string.Empty;
            }

            if (replaceStr == null)
            {
                replaceStr = string.Empty;
            }

            var options = RegexOptions.CultureInvariant;
            if (isIgnoreCase)
            {
                options |= RegexOptions.IgnoreCase;
            }

            if (isMultiline)
            {
                options |= RegexOptions.Multiline;
            }

            var regex = new Regex(matchPattern, options);
            if (isGlobal)
            {
                return regex.Replace(str, replaceStr);
            }

            return regex.Replace(str, replaceStr, 1);
        }

        public bool StartsWith(string str, object prefixes)
        {
            if (str == null)
            {
                str = string.Empty;
            }
            var values = NormalizeStringList(prefixes, "prefixes");
            foreach (var prefix in values)
            {
                if (!string.IsNullOrEmpty(prefix) && str.StartsWith(prefix, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        public bool EndsWith(string str, object suffixes)
        {
            if (str == null)
            {
                str = string.Empty;
            }
            var values = NormalizeStringList(suffixes, "suffixes");
            foreach (var suffix in values)
            {
                if (!string.IsNullOrEmpty(suffix) && str.EndsWith(suffix, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        public long GetWorkbookIndex(Excel.Workbook targetWorkbook)
        {
            if (targetWorkbook == null)
            {
                return 0;
            }

            try
            {
                int count = _app.Workbooks.Count;
                for (int i = 1; i <= count; i++)
                {
                    Excel.Workbook wb = null;
                    try
                    {
                        wb = _app.Workbooks[i];
                        if (wb != null && string.Equals(wb.FullName, targetWorkbook.FullName, StringComparison.OrdinalIgnoreCase))
                        {
                            return i;
                        }
                    }
                    finally
                    {
                        ReleaseCom(wb);
                    }
                }
            }
            catch
            {
            }

            return 0;
        }

        public bool IsSheetExists(string targetSheetName)
        {
            if (string.IsNullOrEmpty(targetSheetName))
            {
                return false;
            }

            try
            {
                foreach (object sheetObj in _app.Worksheets)
                {
                    if (sheetObj is Excel.Worksheet sheet)
                    {
                        try
                        {
                            if (string.Equals(sheet.Name, targetSheetName, StringComparison.Ordinal))
                            {
                                return true;
                            }
                        }
                        finally
                        {
                            ReleaseCom(sheet);
                        }
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        public long GetVisibleSheetsCount()
        {
            long count = 0;
            try
            {
                foreach (object sheetObj in _app.Sheets)
                {
                    if (sheetObj is Excel.Worksheet sheet)
                    {
                        try
                        {
                            if (sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                            {
                                count++;
                            }
                        }
                        finally
                        {
                            ReleaseCom(sheet);
                        }
                    }
                }
            }
            catch
            {
            }

            return count;
        }

        public object[] DirGrob(string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath))
            {
                return Array.Empty<object>();
            }

            folderPath = folderPath.Replace("/", "\\");
            int sepIndex = folderPath.LastIndexOf("\\", StringComparison.Ordinal);
            string lastPart = sepIndex >= 0 ? folderPath.Substring(sepIndex + 1) : folderPath;
            string basePath = sepIndex >= 0 ? folderPath.Substring(0, sepIndex) : string.Empty;

            var results = new List<object>();
            if (Directory.Exists(basePath))
            {
                foreach (var dir in Directory.GetDirectories(basePath))
                {
                    var name = Path.GetFileName(dir);
                    if (!string.IsNullOrEmpty(name) && name.StartsWith(lastPart, StringComparison.OrdinalIgnoreCase))
                    {
                        results.Add(name + "/");
                    }
                }

                foreach (var file in Directory.GetFiles(basePath))
                {
                    var name = Path.GetFileName(file);
                    if (!string.IsNullOrEmpty(name) && name.StartsWith(lastPart, StringComparison.OrdinalIgnoreCase))
                    {
                        results.Add(name);
                    }
                }
            }

            return results.ToArray();
        }

        public string GetAbsolutePath(string cwd, string relativePath)
        {
            if (cwd == null)
            {
                cwd = string.Empty;
            }

            if (relativePath == null)
            {
                relativePath = string.Empty;
            }

            string fullPath;
            if (string.IsNullOrEmpty(cwd))
            {
                fullPath = Path.Combine(Path.DirectorySeparatorChar.ToString(), relativePath);
            }
            else
            {
                fullPath = Path.Combine(cwd, relativePath);
            }

            return Path.GetFullPath(fullPath);
        }

        public string ResolvePath(string strPath)
        {
            if (strPath == null)
            {
                strPath = string.Empty;
            }
            strPath = strPath.Replace("/", "\\");

            string absPath;
            if (strPath.StartsWith("\\", StringComparison.Ordinal))
            {
                absPath = GetAbsolutePath(string.Empty, strPath.Substring(1));
            }
            else if (strPath.StartsWith("~\\", StringComparison.Ordinal))
            {
                var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                absPath = GetAbsolutePath(profile, strPath.Substring(2));
            }
            else
            {
                string basePath = string.Empty;
                try
                {
                    var wb = _app.ActiveWorkbook;
                    basePath = wb?.Path ?? string.Empty;
                }
                catch
                {
                    basePath = string.Empty;
                }

                absPath = GetAbsolutePath(basePath, strPath);
            }

            if (strPath.EndsWith("\\", StringComparison.Ordinal) && !absPath.EndsWith("\\", StringComparison.Ordinal))
            {
                absPath += "\\";
            }

            return absPath;
        }

        public long HexColorCodeToLong(string colorCode)
        {
            if (string.IsNullOrEmpty(colorCode))
            {
                return -1;
            }

            foreach (char ch in colorCode)
            {
                if (!Uri.IsHexDigit(ch))
                {
                    return -1;
                }
            }

            if (colorCode.Length == 3)
            {
                string hex = string.Concat(
                    colorCode[2], colorCode[2],
                    colorCode[1], colorCode[1],
                    colorCode[0], colorCode[0]);
                return Convert.ToInt32(hex, 16);
            }

            if (colorCode.Length == 6)
            {
                string hex = string.Concat(
                    colorCode[4], colorCode[5],
                    colorCode[2], colorCode[3],
                    colorCode[0], colorCode[1]);
                return Convert.ToInt32(hex, 16);
            }

            return -1;
        }

        public string ColorCodeToHex(long colorCode)
        {
            string part1 = (colorCode % 256).ToString("X2", CultureInfo.InvariantCulture);
            string part2 = ((colorCode / 256) % 256).ToString("X2", CultureInfo.InvariantCulture);
            string part3 = ((colorCode / 65536) % 256).ToString("X2", CultureInfo.InvariantCulture);
            return (part1 + part2 + part3).ToLowerInvariant();
        }

        public bool IsJisKeyboardLayout()
        {
            var buffer = new StringBuilder(10);
            if (!GetKeyboardLayoutName(buffer))
            {
                return false;
            }

            string layout = buffer.ToString().TrimEnd('\0');
            return layout.EndsWith("0411", StringComparison.OrdinalIgnoreCase);
        }

        public Excel.Range Union2(object argList)
        {
            var ranges = ExtractRangesList(argList);
            Excel.Range result = null;

            foreach (var range in ranges)
            {
                if (result == null)
                {
                    result = range;
                }
                else
                {
                    result = _app.Union(result, range);
                }
            }

            return result;
        }

        public Excel.Range Intersect2(object argList)
        {
            var ranges = ExtractRangesList(argList, true, out bool invalid);
            if (invalid)
            {
                return null;
            }

            Excel.Range result = null;
            foreach (var range in ranges)
            {
                if (result == null)
                {
                    result = range;
                }
                else
                {
                    result = _app.Intersect(result, range);
                    if (!RangeHelpers.IsRangeValid(result))
                    {
                        return null;
                    }
                }
            }

            return result;
        }

        public Excel.Range Except2(object sourceRange, object argList)
        {
            if (!(sourceRange is Excel.Range source))
            {
                return null;
            }

            Excel.Range buf = source;
            var ranges = ExtractRangesList(argList);
            foreach (var range in ranges)
            {
                var inverted = Invert2(range);
                if (inverted != null)
                {
                    buf = Intersect2(new object[] { buf, inverted });
                    ReleaseCom(inverted);
                    if (!RangeHelpers.IsRangeValid(buf))
                    {
                        return null;
                    }
                }
            }

            return buf;
        }

        public Excel.Range Invert2(object sourceRange)
        {
            if (!(sourceRange is Excel.Range source))
            {
                return null;
            }

            Excel.Worksheet sheet = null;
            Excel.Range buffer = null;

            try
            {
                sheet = source.Parent as Excel.Worksheet;
                if (sheet == null)
                {
                    return null;
                }

                int rowCount = sheet.Rows.Count;
                int colCount = sheet.Columns.Count;
                buffer = sheet.Cells;

                foreach (Excel.Range area in source.Areas)
                {
                    try
                    {
                        int areaTop = area.Row;
                        int areaBottom = areaTop + area.Rows.Count - 1;
                        int areaLeft = area.Column;
                        int areaRight = areaLeft + area.Columns.Count - 1;

                        var rangeLeft = GetRangeWithPosition(sheet, 1, 1, rowCount, areaLeft - 1);
                        var rangeRight = GetRangeWithPosition(sheet, 1, areaRight + 1, rowCount, colCount);
                        var rangeTop = GetRangeWithPosition(sheet, 1, areaLeft, areaTop - 1, areaRight);
                        var rangeBottom = GetRangeWithPosition(sheet, areaBottom + 1, areaLeft, rowCount, areaRight);

                        var union = Union2(new object[] { rangeLeft, rangeRight, rangeTop, rangeBottom });
                        buffer = Intersect2(new object[] { buffer, union });

                        ReleaseCom(rangeLeft);
                        ReleaseCom(rangeRight);
                        ReleaseCom(rangeTop);
                        ReleaseCom(rangeBottom);
                        ReleaseCom(union);
                    }
                    finally
                    {
                        ReleaseCom(area);
                    }
                }
            }
            catch
            {
                ReleaseCom(buffer);
                buffer = null;
            }

            return buffer;
        }

        public bool IsRangeValid(Excel.Range candidate)
            => RangeHelpers.IsRangeValid(candidate);

        public void DebugPrint(string message, string funcName, bool debugMode, string statusPrefix)
        {
            if (!debugMode)
            {
                return;
            }

            string prefix = string.IsNullOrEmpty(funcName) ? string.Empty : "[" + funcName + "] ";
            string text = "[DEBUG] " + prefix + (message ?? string.Empty);

            SetStatusBarTemporarily(text, 5000, false, statusPrefix);
            Debug.WriteLine("[" + DateTime.Now + "] " + text);
        }

        public bool ErrorHandler(int errNumber, string errDescription, string funcName, string statusPrefix)
        {
            if (errNumber == 0)
            {
                return false;
            }

            string text = "[ERROR] ";
            if (!string.IsNullOrEmpty(funcName))
            {
                text += funcName + ": ";
            }

            text += (errDescription ?? string.Empty) + " (" + errNumber + ")";
            SetStatusBarTemporarily(text, 5000, false, statusPrefix);
            Debug.WriteLine("[" + DateTime.Now + "] " + text);
            return true;
        }

        private static IEnumerable<string> NormalizeStringList(object value, string argName)
        {
            if (value == null)
            {
                return Array.Empty<string>();
            }

            if (value is string single)
            {
                return new[] { single };
            }

            if (value is Array array)
            {
                var list = new List<string>();
                foreach (var item in array)
                {
                    if (item is string str)
                    {
                        list.Add(str);
                    }
                }

                return list;
            }

            throw new ArgumentException("Type mismatch: '" + argName + "' must be either a String or String()");
        }

        private List<Excel.Range> ExtractRangesList(object argList)
        {
            bool unused;
            return ExtractRangesList(argList, false, out unused);
        }

        private List<Excel.Range> ExtractRangesList(object argList, bool strict, out bool anyNonRange)
        {
            anyNonRange = false;
            var ranges = new List<Excel.Range>();

            if (argList == null)
            {
                return ranges;
            }

            if (argList is Excel.Range range)
            {
                ranges.Add(range);
                return ranges;
            }

            if (argList is Array array)
            {
                foreach (var item in array)
                {
                    if (item is Excel.Range itemRange)
                    {
                        ranges.Add(itemRange);
                    }
                    else if (strict)
                    {
                        anyNonRange = true;
                    }
                }

                return ranges;
            }

            if (strict)
            {
                anyNonRange = true;
            }

            return ranges;
        }

        private static Excel.Range GetRangeWithPosition(Excel.Worksheet sheet, int top, int left, int bottom, int right)
        {
            if (sheet == null)
            {
                return null;
            }

            if (top > bottom || left > right)
            {
                return null;
            }

            if (top < 1 || left < 1)
            {
                return null;
            }

            try
            {
                if (bottom > sheet.Rows.Count || right > sheet.Columns.Count)
                {
                    return null;
                }

                return sheet.Range[sheet.Cells[top, left], sheet.Cells[bottom, right]];
            }
            catch
            {
                return null;
            }
        }

        private static int GetDecimalDigits(string format)
        {
            int dotIndex = format.IndexOf('.');
            if (dotIndex < 0)
            {
                return 0;
            }

            int count = 0;
            for (int i = dotIndex + 1; i < format.Length; i++)
            {
                char ch = format[i];
                if (ch == '0' || ch == '#')
                {
                    count++;
                }
                else
                {
                    break;
                }
            }

            return count;
        }

        private static double RoundDown(double value, int decimals)
        {
            double scale = Math.Pow(10, Math.Max(0, decimals));
            return Math.Floor(value * scale) / scale;
        }

        private static bool ShouldUpdateStatus(long lastTicks)
        {
            if (lastTicks == 0)
            {
                return true;
            }

            double elapsed = (Stopwatch.GetTimestamp() - lastTicks) / (double)Stopwatch.Frequency;
            return elapsed > 0.1;
        }

        private bool ShouldUpdateStatus()
            => ShouldUpdateStatus(_lastStatusTicks);

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
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern bool GetKeyboardLayoutName(StringBuilder pwszKLID);

        private void OnStatusTimerTick(object sender, EventArgs e)
        {
            _statusTimer.Stop();
            _tempVisible = false;

            try
            {
                if (string.IsNullOrEmpty(_savedStatusMessage))
                {
                    _app.StatusBar = false;
                }
                else
                {
                    _app.StatusBar = _savedStatusMessage;
                }
            }
            catch
            {
            }
        }

        private string BuildProgressText(double percent, int numDigitsAfterDecimal, long currentCount, long maximumCount)
        {
            if (percent < 0)
            {
                percent = 0;
            }
            else if (percent > 100)
            {
                percent = 100;
            }

            int filled = (int)Math.Floor(percent * (ProgressBarLength / 100.0));
            if (filled < 0)
            {
                filled = 0;
            }
            else if (filled > ProgressBarLength)
            {
                filled = ProgressBarLength;
            }

            string bar = new string('#', filled) + new string('-', ProgressBarLength - filled);
            string format = "0" + (numDigitsAfterDecimal > 0 ? "." + new string('0', numDigitsAfterDecimal) : string.Empty);
            double rounded = RoundDown(percent, numDigitsAfterDecimal);
            string percentText = rounded.ToString(format, CultureInfo.CurrentCulture);

            string text = "        Progress:[" + bar + "] " + percentText + " %";
            if (currentCount >= 0 && maximumCount >= currentCount)
            {
                text += " ( " + currentCount + " / " + maximumCount + " )";
            }

            return text;
        }
    }
}
