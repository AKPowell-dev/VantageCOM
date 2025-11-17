using System;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace VantagePackageHolder
{
    internal sealed class PowerPointExporter
    {
        public void PasteClipboardIntoActiveSlide()
        {
            PowerPoint.Application pptApp = null;
            try
            {
                pptApp = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch (COMException)
            {
                System.Windows.Forms.MessageBox.Show(
                    "PowerPoint is not running. Please open a presentation.",
                    "Copy to PowerPoint",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            if (pptApp == null || pptApp.ActiveWindow == null)
            {
                System.Windows.Forms.MessageBox.Show(
                    "No active PowerPoint window detected. Please select a slide and try again.",
                    "Copy to PowerPoint",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            var slide = pptApp.ActiveWindow.View?.Slide as PowerPoint.Slide;
            if (slide == null)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Please make sure a slide is selected in PowerPoint before running this command.",
                    "Copy to PowerPoint",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            try
            {
                var pasted = slide.Shapes.Paste();
                if (pasted != null)
                {
                    pasted.Select();
                    pptApp.Activate();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error pasting into PowerPoint: " + ex.Message,
                    "Copy to PowerPoint",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
