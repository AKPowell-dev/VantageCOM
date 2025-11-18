using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace VantagePackageHolder
{
    internal sealed class PowerPointExporter
    {
        public void PasteClipboardIntoActiveSlide()
        {
            PowerPoint.Application pptApp;
            try
            {
                pptApp = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch (COMException)
            {
                MessageBox.Show(
                    "PowerPoint is not running. Please open a presentation.",
                    "Copy to PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                return;
            }

            var window = pptApp?.ActiveWindow;
            if (window == null)
            {
                MessageBox.Show(
                    "No active PowerPoint window detected. Please select a slide and try again.",
                    "Copy to PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                return;
            }

            var slide = window.View?.Slide as PowerPoint.Slide;
            if (slide == null)
            {
                MessageBox.Show(
                    "Please make sure a slide is selected in PowerPoint before running this command.",
                    "Copy to PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                return;
            }

            PowerPoint.ShapeRange targetRange = null;
            try
            {
                targetRange = window.Selection?.ShapeRange;
            }
            catch
            {
                targetRange = null;
            }

            bool hadTarget = targetRange != null && targetRange.Count > 0;
            double targetLeft = 0, targetTop = 0, targetWidth = 0, targetHeight = 0;
            int desiredZ = 0;

            if (hadTarget)
            {
                try
                {
                    var reference = targetRange.Count >= 1 ? targetRange[1] : null;
                    if (reference != null)
                    {
                        targetLeft = reference.Left;
                        targetTop = reference.Top;
                        targetWidth = reference.Width;
                        targetHeight = reference.Height;
                        desiredZ = reference.ZOrderPosition;
                    }
                }
                catch
                {
                    hadTarget = false;
                }
            }

            PowerPoint.Shape pastedShape;
            try
            {
                pastedShape = PasteShape(slide);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error pasting into PowerPoint: " + ex.Message,
                    "Copy to PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (pastedShape == null)
            {
                MessageBox.Show(
                    "Unable to paste picture into PowerPoint.",
                    "Copy to PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (hadTarget)
            {
                ApplyReplacementTransform(slide, pastedShape, targetRange, targetLeft, targetTop, targetWidth, targetHeight, desiredZ);
            }
            else
            {
                CenterShapeOnSlide(slide, pastedShape);
            }

            try { pastedShape.Select(); } catch { }
            try { pptApp.Activate(); } catch { }
        }

        private PowerPoint.Shape PasteShape(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return null;
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
                return null;
            }

            if (pastedObj is PowerPoint.ShapeRange range)
            {
                try
                {
                    if (range.Count > 0)
                    {
                        return range[1];
                    }
                }
                catch
                {
                    // ignore and fall back
                }
            }

            return pastedObj as PowerPoint.Shape;
        }

        private void ApplyReplacementTransform(
            PowerPoint.Slide slide,
            PowerPoint.Shape pastedShape,
            PowerPoint.ShapeRange targetRange,
            double targetLeft,
            double targetTop,
            double targetWidth,
            double targetHeight,
            int desiredZ)
        {
            double scaledWidth = pastedShape.Width;
            double scaledHeight = pastedShape.Height;

            if (pastedShape.Width > 0 && pastedShape.Height > 0)
            {
                double scaleFactor = 1;
                if (targetWidth > 0 && targetHeight > 0)
                {
                    double scaleX = targetWidth / pastedShape.Width;
                    double scaleY = targetHeight / pastedShape.Height;
                    scaleFactor = Math.Min(scaleX, scaleY);
                }
                else if (targetWidth > 0)
                {
                    scaleFactor = targetWidth / pastedShape.Width;
                }
                else if (targetHeight > 0)
                {
                    scaleFactor = targetHeight / pastedShape.Height;
                }

                if (scaleFactor > 0)
                {
                    scaledWidth = pastedShape.Width * scaleFactor;
                    scaledHeight = pastedShape.Height * scaleFactor;
                }
            }

            pastedShape.Width = (float)scaledWidth;
            pastedShape.Height = (float)scaledHeight;

            double offsetLeft = targetWidth > 0 ? targetLeft + (targetWidth - scaledWidth) / 2.0 : targetLeft;
            double offsetTop = targetHeight > 0 ? targetTop + (targetHeight - scaledHeight) / 2.0 : targetTop;

            pastedShape.Left = (float)offsetLeft;
            pastedShape.Top = (float)offsetTop;

            TryDeleteShapeRange(targetRange);

            if (desiredZ > 0)
            {
                if (desiredZ > slide.Shapes.Count)
                {
                    desiredZ = slide.Shapes.Count;
                }

                try { pastedShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack); } catch { }
                for (int i = 1; i < desiredZ; i++)
                {
                    try { pastedShape.ZOrder(Office.MsoZOrderCmd.msoBringForward); } catch { break; }
                }
            }
        }

        private void CenterShapeOnSlide(PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            double slideWidth = GetSlideWidth(slide);
            double slideHeight = GetSlideHeight(slide);

            shape.Left = (float)((slideWidth - shape.Width) / 2.0);
            shape.Top = (float)((slideHeight - shape.Height) / 2.0);
        }

        private void TryDeleteShapeRange(PowerPoint.ShapeRange range)
        {
            if (range == null)
            {
                return;
            }

            try
            {
                range.Delete();
            }
            catch
            {
                // ignore
            }
        }

        private double GetSlideWidth(PowerPoint.Slide slide)
        {
            try
            {
                return slide.Master?.Width
                    ?? slide.Application?.ActivePresentation?.PageSetup?.SlideWidth
                    ?? 960.0;
            }
            catch
            {
                return 960.0;
            }
        }

        private double GetSlideHeight(PowerPoint.Slide slide)
        {
            try
            {
                return slide.Master?.Height
                    ?? slide.Application?.ActivePresentation?.PageSetup?.SlideHeight
                    ?? 540.0;
            }
            catch
            {
                return 540.0;
            }
        }
    }
}
