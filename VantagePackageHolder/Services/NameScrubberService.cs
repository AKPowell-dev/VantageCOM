using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class NameScrubberService
    {
        private const int LargeNameCount = 25000;

        private readonly Excel.Application _app;
        private NameScrubberForm _form;

        public NameScrubberService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public void ShowDialog()
        {
            if (_app == null)
            {
                return;
            }

            Excel.Workbook wb = null;
            try
            {
                if (_app.Workbooks.Count == 0)
                {
                    return;
                }
                wb = _app.ActiveWorkbook;
            }
            catch
            {
                return;
            }

            if (wb == null)
            {
                return;
            }

            int nameCount = 0;
            try
            {
                nameCount = wb.Names.Count;
            }
            catch
            {
                nameCount = 0;
            }

            if (nameCount <= 0)
            {
                MessageBox.Show("No named ranges were found in the active workbook.",
                    "Name Scrubber",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            if (IsInEditMode(_app))
            {
                MessageBox.Show("Exit cell edit mode before running Name Scrubber.",
                    "Name Scrubber",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (nameCount > LargeNameCount && !IsWorkbookSaved(wb))
            {
                MessageBox.Show("Please save the workbook before running Name Scrubber on large name lists.",
                    "Name Scrubber",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (_form == null || _form.IsDisposed)
            {
                _form = new NameScrubberForm(_app);
                _form.FormClosed += (_, __) => _form = null;
            }

            if (!_form.Visible)
            {
                _form.Show();
            }
            _form.BringToFront();
            _form.Activate();
        }

        private static bool IsInEditMode(Excel.Application app)
        {
            try
            {
                return !app.Ready;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsWorkbookSaved(Excel.Workbook wb)
        {
            try
            {
                return wb.Saved;
            }
            catch
            {
                return false;
            }
        }
    }
}
