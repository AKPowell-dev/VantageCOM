using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class CleanNamesOptions
    {
        public bool DeepClean { get; set; }
        public bool IncludeHidden { get; set; }
    }

    internal sealed class CleanNamesDialog : Form
    {
        private readonly Excel.Workbook _workbook;
        private readonly RadioButton _basicOption;
        private readonly RadioButton _deepOption;
        private readonly CheckBox _includeHidden;
        private readonly Button _saveButton;
        private readonly Button _okButton;
        private readonly Button _cancelButton;

        public CleanNamesOptions Options { get; private set; }

        public CleanNamesDialog(Excel.Workbook workbook)
        {
            _workbook = workbook;

            Text = "Clean Names";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Size = new Size(360, 220);

            _basicOption = new RadioButton { Text = "Basic (errors only)", Checked = true, AutoSize = true };
            _deepOption = new RadioButton { Text = "Deep (errors + missing links)", AutoSize = true };
            _includeHidden = new CheckBox { Text = "Include hidden names", AutoSize = true };
            _saveButton = new Button { Text = "Save Workbook" };
            _okButton = new Button { Text = "OK", DialogResult = DialogResult.OK };
            _cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };

            _saveButton.Click += (_, __) =>
            {
                try
                {
                    _workbook?.Save();
                }
                catch
                {
                    // ignore
                }
            };

            var optionsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                Padding = new Padding(10, 10, 10, 10)
            };
            optionsPanel.Controls.Add(_basicOption);
            optionsPanel.Controls.Add(_deepOption);
            optionsPanel.Controls.Add(_includeHidden);
            optionsPanel.Controls.Add(_saveButton);

            var buttonsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = true,
                Padding = new Padding(10, 6, 10, 10)
            };
            buttonsPanel.Controls.Add(_cancelButton);
            buttonsPanel.Controls.Add(_okButton);

            Controls.Add(buttonsPanel);
            Controls.Add(optionsPanel);

            AcceptButton = _okButton;
            CancelButton = _cancelButton;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                Options = new CleanNamesOptions
                {
                    DeepClean = _deepOption.Checked,
                    IncludeHidden = _includeHidden.Checked
                };
            }

            base.OnFormClosing(e);
        }
    }

    internal sealed class GoToForm : Form
    {
        private readonly Excel.Application _app;
        private readonly List<Excel.Range> _ranges;
        private readonly ListBox _listBox;
        private readonly Button _closeButton;

        public GoToForm(Excel.Application app, List<Excel.Range> ranges)
        {
            _app = app;
            _ranges = ranges ?? new List<Excel.Range>();

            Text = "Go To";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            Size = new Size(360, 320);
            ShowInTaskbar = false;

            _listBox = new ListBox { Dock = DockStyle.Fill };
            _closeButton = new Button { Text = "Close", Dock = DockStyle.Bottom };
            _closeButton.Click += (_, __) => Close();
            _listBox.SelectedIndexChanged += (_, __) => NavigateToSelected();

            Controls.Add(_listBox);
            Controls.Add(_closeButton);

            Load += (_, __) =>
            {
                _listBox.Items.Clear();
                foreach (var range in _ranges)
                {
                    _listBox.Items.Add(FormatRange(range));
                }
            };
        }

        private void NavigateToSelected()
        {
            int index = _listBox.SelectedIndex;
            if (index < 0 || index >= _ranges.Count)
            {
                return;
            }

            var range = _ranges[index];
            if (range == null)
            {
                return;
            }

            try
            {
                RangeHelpers.SafeActivateSheet(range.Worksheet);
                _app.Goto(range, true);
            }
            catch
            {
                // ignore
            }
        }

        private static string FormatRange(Excel.Range range)
        {
            if (range == null)
            {
                return string.Empty;
            }

            try
            {
                var sheet = range.Worksheet?.Name ?? string.Empty;
                var address = range.Address[false, false, Excel.XlReferenceStyle.xlA1];
                return string.IsNullOrEmpty(sheet) ? address : sheet + "!" + address;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
