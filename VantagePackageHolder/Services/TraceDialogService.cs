using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class TraceDialogService
    {
        private readonly Excel.Application _app;
        private TraceDialogForm _precedentsForm;
        private TraceDialogForm _dependentsForm;

        public TraceDialogService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public void ShowPrecedentsDialog() => ShowDialog(TraceDialogMode.Precedents);

        public void ShowDependentsDialog() => ShowDialog(TraceDialogMode.Dependents);

        private void ShowDialog(TraceDialogMode mode)
        {
            var form = mode == TraceDialogMode.Precedents ? _precedentsForm : _dependentsForm;
            if (mode == TraceDialogMode.Dependents && form != null && !form.IsDisposed)
            {
                try
                {
                    form.Close();
                }
                catch
                {
                    // ignore
                }
                form = null;
                _dependentsForm = null;
            }
            if (form == null || form.IsDisposed)
            {
                form = new TraceDialogForm(_app, mode);
                form.FormClosed += (_, __) =>
                {
                    if (mode == TraceDialogMode.Precedents)
                    {
                        _precedentsForm = null;
                    }
                    else
                    {
                        _dependentsForm = null;
                    }
                };

                if (mode == TraceDialogMode.Precedents)
                {
                    _precedentsForm = form;
                }
                else
                {
                    _dependentsForm = form;
                }
            }

            try
            {
                form.RefreshFromSelection();
            }
            catch
            {
                form.RefreshFromSelection();
            }

            if (!form.HasContent)
            {
                return;
            }

            if (!form.Visible)
            {
                form.Show();
            }
            form.BringToFront();
            form.Activate();
        }
    }
}
