using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
namespace VantagePackageHolder
{
    internal sealed class BatchResizeService
    {
        private const int MaxAttempts = 6;
        private const double TolerancePoints = 0.02;
        private const double MinAdjustColumnWidth = 0.0;
        private const double MeasurementOutlierRatio = 1.25;
        private const double CalibrationMinRatio = 0.5;
        private const double CalibrationMaxRatio = 1.5;

        private readonly Excel.Application _app;
        private readonly FormatService _format;
        private double? _emfCalibrationFactor;

        public BatchResizeService(Excel.Application app, FormatService format)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _format = format ?? throw new ArgumentNullException(nameof(format));
        }

        public void ResizeSelectionToWidthInches(double targetInches, bool requirePowerPoint)
        {
            if (targetInches <= 0)
            {
                return;
            }

            Excel.Range selection = null;
            try
            {
                selection = _app.Selection as Excel.Range;
            }
            catch
            {
                return;
            }

            if (!RangeHelpers.IsRangeValid(selection))
            {
                return;
            }

            double targetWidthPts = targetInches * 72.0;

            try
            {
                using (new UiGuard(_app, hideStatusBar: true))
                {
                    if (requirePowerPoint)
                    {
                        ResizeSelectionByMeasuredWidth(selection, targetWidthPts);
                        return;
                    }

                    for (int attempt = 0; attempt < MaxAttempts; attempt++)
                    {
                        double currentWidthPts = selection.Width * GetPrintScaleFactor(selection.Worksheet);

                        if (currentWidthPts <= 0)
                        {
                            return;
                        }

                        if (Math.Abs(currentWidthPts - targetWidthPts) <= TolerancePoints)
                        {
                            break;
                        }

                        GetColumnWidthBuckets(selection, MinAdjustColumnWidth, out double fixedPts, out double adjustablePts);
                        double totalExcelPts = fixedPts + adjustablePts;
                        if (totalExcelPts <= 0)
                        {
                            return;
                        }

                        double k = currentWidthPts / totalExcelPts;
                        if (k <= 0)
                        {
                            return;
                        }

                        double desiredExcelPts = targetWidthPts / k;
                        if (adjustablePts <= 0 || desiredExcelPts <= fixedPts)
                        {
                            return;
                        }

                        double scaleFactor = (desiredExcelPts - fixedPts) / adjustablePts;
                        ScaleSelectionColumns(selection, scaleFactor, MinAdjustColumnWidth);
                    }
                }
            }
            finally
            {
            }
        }

        private double MeasureCopyPictureWidthPts(Excel.Range selection)
        {
            if (selection == null)
            {
                return 0;
            }

            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                return 0;
            }

            double expectedWidthPts = 0;
            try
            {
                expectedWidthPts = selection.Width * GetPrintScaleFactor(selection.Worksheet);
            }
            catch
            {
                expectedWidthPts = 0;
            }

            double[] samples = new double[3];
            int sampleCount = 0;

            for (int attempt = 0; attempt < 3; attempt++)
            {
                ClearClipboardBestEffort();

                try
                {
                    if (!_format.CopySelectionAsPicturePrintSafe())
                    {
                        continue;
                    }

                    double widthPts = WaitForClipboardEmfWidthPts(650);
                    if (widthPts <= 0)
                    {
                        widthPts = TryGetClipboardPictureWidthPts();
                    }

                    if (widthPts > 0)
                    {
                        samples[sampleCount++] = widthPts;
                    }
                }
                catch
                {
                }

                Thread.Sleep(20);
            }

            if (sampleCount == 0)
            {
                if (expectedWidthPts > 0 && _emfCalibrationFactor.HasValue)
                {
                    return expectedWidthPts * _emfCalibrationFactor.Value;
                }

                return 0;
            }

            Array.Sort(samples, 0, sampleCount);
            double measured = samples[sampleCount / 2];

            if (expectedWidthPts > 0)
            {
                double ratio = measured / expectedWidthPts;
                if (ratio >= CalibrationMinRatio && ratio <= CalibrationMaxRatio)
                {
                    _emfCalibrationFactor = ratio;
                }

                if (IsMeasurementOutlier(measured, expectedWidthPts) && _emfCalibrationFactor.HasValue)
                {
                    return expectedWidthPts * _emfCalibrationFactor.Value;
                }
            }

            return measured;
        }

        private static double WaitForClipboardEmfWidthPts(int maxMillis)
        {
            int deadline = Environment.TickCount + Math.Max(60, maxMillis);
            while (Environment.TickCount < deadline)
            {
                double widthPts = TryGetClipboardEmfWidthPtsViaWin32();
                if (widthPts > 0)
                {
                    return widthPts;
                }

                Application.DoEvents();
                Thread.Sleep(10);
            }

            return 0;
        }

        private static bool IsMeasurementOutlier(double measured, double expected)
        {
            if (expected <= 0 || measured <= 0)
            {
                return false;
            }

            double ratio = measured / expected;
            return ratio > MeasurementOutlierRatio || ratio < (1.0 / MeasurementOutlierRatio);
        }

        private static double TryGetClipboardPictureWidthPts()
        {
            try
            {
                double emfWidth = TryGetClipboardEmfWidthPtsViaWin32();
                if (emfWidth > 0)
                {
                    return emfWidth;
                }

                IDataObject data = Clipboard.GetDataObject();
                if (data != null && data.GetDataPresent(DataFormats.EnhancedMetafile))
                {
                    using (var metafile = data.GetData(DataFormats.EnhancedMetafile) as Metafile)
                    {
                        double widthPts = TryGetMetafileFrameWidthPts(metafile);
                        if (widthPts <= 0)
                        {
                            widthPts = GetWidthPointsFromImage(metafile);
                        }
                        if (widthPts > 0)
                        {
                            return widthPts;
                        }
                    }
                }

                return 0;
            }
            catch
            {
                return 0;
            }
        }

        private static double GetWidthPointsFromImage(Image image)
        {
            if (image == null)
            {
                return 0;
            }

            try
            {
                if (image is Metafile emf)
                {
                    double emfWidth = TryGetMetafileFrameWidthPts(emf);
                    if (emfWidth > 0)
                    {
                        return emfWidth;
                    }

                    GraphicsUnit unit = GraphicsUnit.Inch;
                    RectangleF bounds = emf.GetBounds(ref unit);
                    if (bounds.Width > 0)
                    {
                        return bounds.Width * 72.0;
                    }

                    if (emf.HorizontalResolution > 0)
                    {
                        double widthInches = emf.Width / emf.HorizontalResolution;
                        return widthInches * 72.0;
                    }
                }

                if (image.HorizontalResolution > 0)
                {
                    double widthInches = image.Width / image.HorizontalResolution;
                    return widthInches * 72.0;
                }
            }
            catch
            {
            }

            return 0;
        }

        private static double TryGetMetafileFrameWidthPts(Metafile emf)
        {
            if (emf == null)
            {
                return 0;
            }

            IntPtr hemf = IntPtr.Zero;
            try
            {
                hemf = emf.GetHenhmetafile();
                if (hemf == IntPtr.Zero)
                {
                    return 0;
                }

                if (GetEnhMetaFileHeader(hemf, (uint)Marshal.SizeOf(typeof(ENHMETAHEADER)), out ENHMETAHEADER header) == 0)
                {
                    return 0;
                }

                int frameWidth = header.rclFrame.Right - header.rclFrame.Left;
                if (frameWidth <= 0)
                {
                    return 0;
                }

                double widthInches = frameWidth / 2540.0; // 0.01 mm -> inches
                return widthInches * 72.0;
            }
            catch
            {
                return 0;
            }
            finally
            {
                if (hemf != IntPtr.Zero)
                {
                    try { DeleteEnhMetaFile(hemf); } catch { }
                }
            }
        }

        [DllImport("gdi32.dll")]
        private static extern uint GetEnhMetaFileHeader(IntPtr hemf, uint cbBuffer, out ENHMETAHEADER lpemh);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteEnhMetaFile(IntPtr hemf);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool CloseClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool EmptyClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool IsClipboardFormatAvailable(uint format);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetClipboardData(uint uFormat);

        private const uint CF_ENHMETAFILE = 14;

        private static double TryGetClipboardEmfWidthPtsViaWin32()
        {
            if (!OpenClipboard(IntPtr.Zero))
            {
                return 0;
            }

            try
            {
                if (!IsClipboardFormatAvailable(CF_ENHMETAFILE))
                {
                    return 0;
                }

                IntPtr hemf = GetClipboardData(CF_ENHMETAFILE);
                if (hemf == IntPtr.Zero)
                {
                    return 0;
                }

                if (GetEnhMetaFileHeader(hemf, (uint)Marshal.SizeOf(typeof(ENHMETAHEADER)), out ENHMETAHEADER header) == 0)
                {
                    return 0;
                }

                int frameWidth = header.rclFrame.Right - header.rclFrame.Left;
                if (frameWidth <= 0)
                {
                    return 0;
                }

                double widthInches = frameWidth / 2540.0; // 0.01 mm -> inches
                return widthInches * 72.0;
            }
            catch
            {
                return 0;
            }
            finally
            {
                CloseClipboard();
            }
        }

        private static void ClearClipboardBestEffort()
        {
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    if (OpenClipboard(IntPtr.Zero))
                    {
                        try { EmptyClipboard(); }
                        finally { CloseClipboard(); }
                        break;
                    }
                }
                catch
                {
                }

                Thread.Sleep(10);
            }

            try { Clipboard.Clear(); } catch { }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ENHMETAHEADER
        {
            public uint iType;
            public uint nSize;
            public RECT rclBounds;
            public RECT rclFrame;
            public uint dSignature;
            public uint nVersion;
            public uint nBytes;
            public uint nRecords;
            public ushort nHandles;
            public ushort sReserved;
            public uint nDescription;
            public uint offDescription;
            public uint nPalEntries;
            public SIZE szlDevice;
            public SIZE szlMillimeters;
            public uint cbPixelFormat;
            public uint offPixelFormat;
            public uint bOpenGL;
            public SIZE szlMicrometers;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SIZE
        {
            public int cx;
            public int cy;
        }

        private void ResizeSelectionByMeasuredWidth(Excel.Range selection, double targetWidthPts)
        {
            if (selection == null)
            {
                return;
            }

            Excel.Worksheet ws = null;
            try
            {
                ws = selection.Worksheet;
            }
            catch
            {
                ws = null;
            }

            if (ws == null)
            {
                return;
            }

            System.Collections.Generic.List<ColumnInfo> columns;
            double fixedPts;
            double adjustablePts;
            if (!BuildColumnInfo(selection, MinAdjustColumnWidth, out columns, out fixedPts, out adjustablePts))
            {
                return;
            }

            if (adjustablePts <= 0)
            {
                return;
            }

            double totalExcelPts = fixedPts + adjustablePts;
            if (totalExcelPts <= 0)
            {
                return;
            }

            double measured = MeasureCopyPictureWidthPts(selection);
            if (measured <= 0)
            {
                ResizeSelectionByPrintScale(selection, targetWidthPts, fixedPts, adjustablePts);
                return;
            }

            double scaleFactor = ComputeAdjustableScale(targetWidthPts, measured, fixedPts, adjustablePts, totalExcelPts);
            if (scaleFactor <= 0)
            {
                return;
            }

            ApplyColumnWidthFactor(ws, columns, scaleFactor);

            double measuredAfter = MeasureCopyPictureWidthPts(selection);
            if (measuredAfter <= 0)
            {
                return;
            }

            double delta = Math.Abs(measuredAfter - targetWidthPts);
            if (delta <= TolerancePoints)
            {
                return;
            }

            double correction = targetWidthPts / measuredAfter;
            double correctedFactor = scaleFactor * correction;
            if (correctedFactor > 0)
            {
                ApplyColumnWidthFactor(ws, columns, correctedFactor);
                double measuredAfter2 = MeasureCopyPictureWidthPts(selection);
                if (measuredAfter2 > 0)
                {
                    double delta2 = Math.Abs(measuredAfter2 - targetWidthPts);
                    if (delta2 > TolerancePoints)
                    {
                        double correction2 = targetWidthPts / measuredAfter2;
                        double correctedFactor2 = correctedFactor * correction2;
                        if (correctedFactor2 > 0)
                        {
                            ApplyColumnWidthFactor(ws, columns, correctedFactor2);
                        }
                    }
                }
            }
        }

        private void ResizeSelectionByPrintScale(Excel.Range selection, double targetWidthPts, double fixedPts, double adjustablePts)
        {
            if (selection == null || adjustablePts <= 0)
            {
                return;
            }

            double currentWidthPts = selection.Width * GetPrintScaleFactor(selection.Worksheet);
            if (currentWidthPts <= 0)
            {
                return;
            }

            double totalExcelPts = fixedPts + adjustablePts;
            if (totalExcelPts <= 0)
            {
                return;
            }

            double scaleFactor = ComputeAdjustableScale(targetWidthPts, currentWidthPts, fixedPts, adjustablePts, totalExcelPts);
            if (scaleFactor <= 0)
            {
                return;
            }

            ScaleSelectionColumns(selection, scaleFactor, MinAdjustColumnWidth);
        }

        private static double ComputeAdjustableScale(double targetWidthPts, double measuredWidthPts, double fixedPts, double adjustablePts, double totalExcelPts)
        {
            if (measuredWidthPts <= 0 || adjustablePts <= 0 || totalExcelPts <= 0)
            {
                return 0;
            }

            double k = measuredWidthPts / totalExcelPts;
            if (k <= 0)
            {
                return 0;
            }

            double desiredExcelPts = targetWidthPts / k;
            if (desiredExcelPts <= fixedPts)
            {
                return 0;
            }

            double scaleFactor = (desiredExcelPts - fixedPts) / adjustablePts;
            return scaleFactor;
        }

        private sealed class ColumnInfo
        {
            public int Index { get; set; }
            public double BaseWidth { get; set; }
        }

        private bool BuildColumnInfo(Excel.Range selection, double minAdjustWidth, out System.Collections.Generic.List<ColumnInfo> columns, out double fixedPts, out double adjustablePts)
        {
            columns = new System.Collections.Generic.List<ColumnInfo>();
            fixedPts = 0;
            adjustablePts = 0;

            if (selection == null)
            {
                return false;
            }

            var seen = new System.Collections.Generic.HashSet<int>();

            foreach (Excel.Range column in selection.Columns)
            {
                try
                {
                    if (column.EntireColumn.Hidden)
                    {
                        continue;
                    }

                    int index = column.Column;
                    if (!seen.Add(index))
                    {
                        continue;
                    }

                    double widthPts = column.Width;
                    double widthCol = column.ColumnWidth;
                    if (widthCol <= minAdjustWidth)
                    {
                        fixedPts += widthPts;
                    }
                    else
                    {
                        adjustablePts += widthPts;
                        columns.Add(new ColumnInfo { Index = index, BaseWidth = widthCol });
                    }
                }
                catch
                {
                }
                finally
                {
                    ReleaseCom(column);
                }
            }

            return columns.Count > 0;
        }

        private void ApplyColumnWidthFactor(Excel.Worksheet ws, System.Collections.Generic.List<ColumnInfo> columns, double factor)
        {
            if (ws == null || columns == null || columns.Count == 0 || factor <= 0)
            {
                return;
            }

            foreach (var info in columns)
            {
                Excel.Range column = null;
                try
                {
                    column = ws.Columns[info.Index] as Excel.Range;
                    if (column == null)
                    {
                        continue;
                    }

                    double target = info.BaseWidth * factor;
                    if (target < 0.1)
                    {
                        target = 0.1;
                    }
                    else if (target > 255.0)
                    {
                        target = 255.0;
                    }

                    column.ColumnWidth = target;
                }
                catch
                {
                }
                finally
                {
                    ReleaseCom(column);
                }
            }
        }

        private void ScaleSelectionColumns(Excel.Range selection, double scaleFactor, double minAdjustWidth)
        {
            if (selection == null || scaleFactor <= 0)
            {
                return;
            }

            foreach (Excel.Range column in selection.Columns)
            {
                try
                {
                    if (column.EntireColumn.Hidden)
                    {
                        continue;
                    }

                    if (minAdjustWidth >= 0 && column.ColumnWidth <= minAdjustWidth)
                    {
                        continue;
                    }

                    double widthPts = column.Width;
                    if (widthPts <= 0)
                    {
                        continue;
                    }

                    double widthCol = column.ColumnWidth;
                    double targetWidth = widthCol * scaleFactor;
                    if (targetWidth < 0.1)
                    {
                        targetWidth = 0.1;
                    }
                    else if (targetWidth > 255.0)
                    {
                        targetWidth = 255.0;
                    }

                    column.EntireColumn.ColumnWidth = targetWidth;
                }
                catch
                {
                }
                finally
                {
                    ReleaseCom(column);
                }
            }
        }

        private void GetColumnWidthBuckets(Excel.Range selection, double minAdjustWidth, out double fixedPts, out double adjustablePts)
        {
            fixedPts = 0;
            adjustablePts = 0;

            if (selection == null)
            {
                return;
            }

            foreach (Excel.Range column in selection.Columns)
            {
                try
                {
                    if (column.EntireColumn.Hidden)
                    {
                        continue;
                    }

                    if (column.ColumnWidth <= minAdjustWidth)
                    {
                        fixedPts += column.Width;
                    }
                    else
                    {
                        adjustablePts += column.Width;
                    }
                }
                catch
                {
                }
                finally
                {
                    ReleaseCom(column);
                }
            }
        }

        private static double GetPrintScaleFactor(Excel.Worksheet sheet)
        {
            if (sheet == null)
            {
                return 1.0;
            }

            try
            {
                var zoomValue = sheet.PageSetup.Zoom;
                if (zoomValue is double zoom && zoom > 0)
                {
                    return zoom / 100.0;
                }

                if (zoomValue is int zoomInt && zoomInt > 0)
                {
                    return zoomInt / 100.0;
                }
            }
            catch
            {
            }

            return 1.0;
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
            }
        }
    }
}
