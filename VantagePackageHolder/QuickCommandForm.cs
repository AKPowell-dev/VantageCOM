using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace VantagePackageHolder
{
    internal sealed class QuickCommandForm : Form
    {
        private readonly Label _prefix;
        private readonly TextBox _input;
        public string CommandText => _input.Text;
        private readonly PowerPoint.Slide _slide;

        public QuickCommandForm(PowerPoint.Slide slide)
        {
            _slide = slide;

            _prefix = new Label
            {
                AutoSize = false,
                Text = ":",
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                Width = 18,
                Dock = DockStyle.Left,
                Font = new Font("Consolas", 11f, FontStyle.Regular)
            };

            _input = new TextBox
            {
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Consolas", 11f),
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(20, 20, 20),
                ForeColor = Color.White
            };

            SuspendLayout();
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            ControlBox = false;
            ShowIcon = false;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.FromArgb(32, 32, 32);
            Padding = new Padding(8, 8, 8, 8);
            Size = new Size(360, 48);

            var panel = new Panel
            {
                Dock = DockStyle.Fill
            };
            panel.Controls.Add(_input);
            panel.Controls.Add(_prefix);

            Controls.Add(panel);

            AcceptButton = new DialogCloseButton(DialogResult.OK);
            CancelButton = new DialogCloseButton(DialogResult.Cancel);

            KeyPreview = true;

            Load += (_, __) =>
            {
                // Position at bottom-center of the active slide window if possible
                // Position at bottom-center of the working area (close to slide bottom in normal view)
                Rectangle wa = Screen.PrimaryScreen.WorkingArea;
                int margin = 12;
                Location = new Point(wa.Left + (wa.Width - Width) / 2 + 4, wa.Bottom - Height - margin);
                _input.Focus();
                _input.SelectAll();
            };

            _input.TextChanged += (_, __) =>
            {
                // Keep tabs out and trim width growth a bit like Excel cmdline
                var selStart = _input.SelectionStart;
                _input.Text = _input.Text.Replace("\t", " ");
                _input.SelectionStart = selStart;
            };
            ResumeLayout(performLayout: true);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
                Close();
                return true;
            }
            if (keyData == Keys.Enter || keyData == Keys.Return)
            {
                DialogResult = DialogResult.OK;
                Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private sealed class DialogCloseButton : Button
        {
            public DialogCloseButton(DialogResult result)
            {
                DialogResult = result;
                Size = new Size(0, 0);
                Visible = false;
                TabStop = false;
                CausesValidation = false;
            }
        }
    }
}
