using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace VantagePackageHolder
{
    internal sealed class BatchResizeService
    {
        private const int MaxAttempts = 3;
        private const double TolerancePoints = 0.05;
        private const double MinAdjustColumnWidth = 0.6;

        private readonly Excel.Application _app;
        private readonly FormatService _format;

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

            PowerPoint.Application pptApp = null;
            PowerPoint.DocumentWindow pptWindow = null;
            PowerPoint.Presentation pptPres = null;
            PowerPoint.Slide originalSlide = null;
            PowerPoint.Slide tempSlide = null;
            int originalIndex = 0;

            try
            {
                if (requirePowerPoint && !TryGetPowerPointContext(out pptApp, out pptWindow, out pptPres, out originalSlide))
                {
                    return;
                }

                if (requirePowerPoint)
                {
                    originalIndex = originalSlide?.SlideIndex ?? 0;
                    tempSlide = pptPres?.Slides.Add(pptPres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                    tempSlide?.Select();
                }

                using (new UiGuard(_app, hideStatusBar: true))
                {
                    for (int attempt = 0; attempt < MaxAttempts; attempt++)
                    {
                        double currentWidthPts = requirePowerPoint
                            ? MeasureCopyPictureWidthPts(selection, tempSlide)
                            : selection.Width * GetPrintScaleFactor(selection.Worksheet);

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

                        double scaleFactor;
                        if (adjustablePts <= 0 || desiredExcelPts <= fixedPts)
                        {
                            scaleFactor = targetWidthPts / currentWidthPts;
                            ScaleSelectionColumns(selection, scaleFactor, -1);
                        }
                        else
                        {
                            scaleFactor = (desiredExcelPts - fixedPts) / adjustablePts;
                            ScaleSelectionColumns(selection, scaleFactor, MinAdjustColumnWidth);
                        }
                    }
                }
            }
            finally
            {
                try
                {
                    tempSlide?.Delete();
                }
                catch
                {
                }

                if (pptPres != null && originalIndex > 0)
                {
                    try
                    {
                        pptPres.Slides[originalIndex].Select();
                    }
                    catch
                    {
                    }
                }

                ReleaseCom(tempSlide);
                ReleaseCom(originalSlide);
                ReleaseCom(pptPres);
                ReleaseCom(pptWindow);
                ReleaseCom(pptApp);
            }
        }

        private bool TryGetPowerPointContext(out PowerPoint.Application app, out PowerPoint.DocumentWindow window, out PowerPoint.Presentation presentation, out PowerPoint.Slide slide)
        {
            app = null;
            window = null;
            presentation = null;
            slide = null;

            try
            {
                app = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                window = app?.ActiveWindow;
                presentation = window?.Presentation;
                slide = window?.View?.Slide as PowerPoint.Slide;
            }
            catch
            {
                app = null;
                window = null;
                presentation = null;
                slide = null;
            }

            return app != null && window != null && presentation != null && slide != null;
        }

        private double MeasureCopyPictureWidthPts(Excel.Range selection, PowerPoint.Slide slide)
        {
            if (selection == null || slide == null)
            {
                return 0;
            }

            for (int attempt = 0; attempt < 2; attempt++)
            {
                if (!_format.CopySelectionAsPicturePrintSafe())
                {
                    continue;
                }

                object pastedObj = null;
                try
                {
                    pastedObj = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile);
                }
                catch
                {
                    pastedObj = null;
                }

                if (pastedObj == null)
                {
                    try
                    {
                        pastedObj = slide.Shapes.Paste();
                    }
                    catch
                    {
                        pastedObj = null;
                    }
                }

                if (pastedObj == null)
                {
                    continue;
                }

                double width = 0;
                try
                {
                    PowerPoint.Shape shape = null;
                    if (pastedObj is PowerPoint.ShapeRange range)
                    {
                        if (range.Count > 0)
                        {
                            shape = range[1];
                        }
                    }
                    else
                    {
                        shape = pastedObj as PowerPoint.Shape;
                    }

                    if (shape != null)
                    {
                        width = shape.Width;
                    }

                    try
                    {
                        if (pastedObj is PowerPoint.ShapeRange rangeDelete)
                        {
                            rangeDelete.Delete();
                        }
                        else if (shape != null)
                        {
                            shape.Delete();
                        }
                    }
                    catch
                    {
                    }
                }
                finally
                {
                    ReleaseCom(pastedObj);
                }

                return width;
            }

            return 0;
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
