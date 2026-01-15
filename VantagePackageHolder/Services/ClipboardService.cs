using System;
using System.Windows.Forms; // for Clipboard
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class ClipboardService
    {
        private readonly Excel.Application _app;
        private Excel.Range _copyRange;

        public ClipboardService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public void HandleCopy()
        {
            try
            {
                _copyRange = _app.Selection as Excel.Range;
            }
            catch
            {
                _copyRange = null;
            }

            ExecuteMsoSafe("Copy");
        }

        public void HandleCut()
        {
            _copyRange = null;
            ExecuteMsoSafe("Cut");
        }

        public void HandlePaste()
        {
            EnsureClipboardPayload();
            ExecuteMsoSafe("Paste");
        }

        public void HandlePasteValues()
        {
            EnsureClipboardPayload();
            ExecuteMsoSafe("PasteValues");
        }

        public void HandlePasteFormulas()
        {
            EnsureClipboardPayload();
            ExecuteMsoSafe("PasteFormulas");
        }

        public void PasteValuesSmart()
        {
            using (new UiGuard(_app, hideStatusBar: true))
            {
                var cutCopyMode = _app.CutCopyMode;
                if (cutCopyMode == Excel.XlCutCopyMode.xlCopy || cutCopyMode == Excel.XlCutCopyMode.xlCut)
                {
                    EnsureClipboardPayload();
                    ExecuteMsoSafe("PasteValues");
                    return;
                }

                if (!ClipboardHasContent())
                {
                    return;
                }

                try
                {
                    if (Clipboard.ContainsText(TextDataFormat.Rtf))
                    {
                        if (!TryExecuteMso("PasteKeepTextOnly"))
                        {
                            ExecuteMsoSafe("Paste");
                        }
                        return;
                    }

                    if (Clipboard.ContainsText(TextDataFormat.Html))
                    {
                        ExecuteMsoSafe("PasteSpecialDialog");
                        return;
                    }

                    ExecuteMsoSafe("Paste");
                }
                catch
                {
                    ExecuteMsoSafe("Paste");
                }
            }
        }

        public void OpenPasteSpecialDialog()
        {
            EnsureClipboardPayload();
            ExecuteMsoSafe("PasteSpecial"); // <- correct id
        }

        public Excel.Range GetCopyRange()
            => RangeHelpers.IsRangeValid(_copyRange) ? _copyRange : null;

        public void SetCopyRange(Excel.Range range)
            => _copyRange = RangeHelpers.IsRangeValid(range) ? range : null;

        public bool ClipboardHasContent() => HasClipboardContent();

        private void EnsureClipboardPayload()
        {
            // If the clipboard already has something usable (e.g., a screenshot bitmap), do not overwrite it.
            if (HasClipboardContent())
                return;

            // Skip if Excel is already in copy mode and clipboard has usable content.
            if (_app.CutCopyMode == Excel.XlCutCopyMode.xlCopy && HasClipboardContent())
                return;

            if (!RangeHelpers.IsRangeValid(_copyRange))
            {
                _copyRange = null;
                return;
            }

            Excel.Range savedSelection = null;
            Excel.Range savedActiveCell = null;

            try
            {
                savedSelection = _app.Selection as Excel.Range;
                savedActiveCell = _app.ActiveCell;

                _copyRange.Copy(); // repopulate Excel's CutCopyMode and Windows clipboard
            }
            catch
            {
                _copyRange = null;
            }
            finally
            {
                try
                {
                    if (savedSelection != null &&
                        savedSelection.Worksheet?.Parent == _app.ActiveWorkbook)
                    {
                        RangeHelpers.SafeSelect(savedSelection);
                        try { savedActiveCell?.Activate(); } catch { /* ignore */ }
                    }
                }
                catch { /* ignore */ }
            }
        }

        private bool HasClipboardContent()
        {
            try
            {
                // Check a few common formats Excel can paste
                return Clipboard.ContainsData(DataFormats.UnicodeText)
                    || Clipboard.ContainsData(DataFormats.Text)
                    || Clipboard.ContainsData(DataFormats.EnhancedMetafile)
                    || Clipboard.ContainsData(DataFormats.MetafilePict)
                    || Clipboard.ContainsImage()
                    || Clipboard.ContainsData(DataFormats.Bitmap)
                    || Clipboard.ContainsData(DataFormats.Html)
                    || Clipboard.ContainsData(DataFormats.CommaSeparatedValue);
            }
            catch
            {
                return false;
            }
        }

        private void ExecuteMsoSafe(string controlId) => TryExecuteMso(controlId);

        private bool TryExecuteMso(string controlId)
        {
            try
            {
                _app.CommandBars.ExecuteMso(controlId);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

}
