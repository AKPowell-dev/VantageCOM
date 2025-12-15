using System;
using System.Runtime.InteropServices;
using Extensibility;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Drawing;

namespace VantagePackageHolder
{
    [ComVisible(true)]
    [Guid("1F2E80B3-40C9-4B18-921E-0AB6B63C6C6A")]
    [ProgId("VantagePackageHolder.PowerPointAddin")]
    public sealed class PowerPointAddin : IDTExtensibility2
    {
        private PowerPoint.Application _ppt;
        private PowerPoint.EApplication_Event _events;
        private IntPtr _subclassHandle = IntPtr.Zero;
        private IntPtr _originalWndProc = IntPtr.Zero;
        private WndProcDelegate _wndProcDelegate;
        private QuickCommandForm _activeCmdForm;

        // Hotkey constants
        private const int WM_HOTKEY = 0x0312;
        private const int MOD_CONTROL = 0x0002;
        private const int MOD_SHIFT = 0x0004;
        private const int VK_T = 0x54;
        private const int VK_OEM_1 = 0xBA; // ;/: key
        private const int VK_B = 0x42;
        private const int VK_SHIFT_KEY = 0x10;
        private const int VK_CONTROL_KEY = 0x11;
        private const int VK_MENU_KEY = 0x12;
        private const int HOTKEY_ID = 0xBEEF;
        private const int HOTKEY_SHAPE_ID = 0xBEED;

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _ppt = application as PowerPoint.Application;

            if (addInInst is Office.COMAddIn comAddIn)
            {
                comAddIn.Object = this;
            }

            try
            {
                _events = (PowerPoint.EApplication_Event)_ppt;
            }
            catch
            {
                // Ignore event hookup failures; add-in still functions.
            }

            InstallHotkey();
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                _events = null;
            }
            catch
            {
                // ignore
            }

            UninstallHotkey();
            TearDownCmdForm();
            _ppt = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        /// <summary>
        /// Inserts a simple textbox on the active slide to verify the add-in is loaded.
        /// </summary>
        public void InsertDemoTextbox()
        {
            if (_ppt == null) return;

            PowerPoint.Slide targetSlide = ResolveActiveSlide();
            if (targetSlide == null) return;

            try
            {
                const float width = 320f;
                const float height = 80f;
                const float left = 120f;
                const float top = 120f;

                var shape = targetSlide.Shapes.AddTextbox(
                    Orientation: Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    Left: left,
                    Top: top,
                    Width: width,
                    Height: height);

                var text = shape.TextFrame.TextRange;
                text.Text = "Vantage add-in test box";
                text.Font.Size = 20;
                text.Font.Bold = Office.MsoTriState.msoTrue;
            }
            catch
            {
                // Swallow errors to avoid crashing PowerPoint.
            }
        }

        /// <summary>
        /// Inserts a textbox with Garamond 11 when the hotkey is triggered.
        /// </summary>
        public void InsertStyledTextbox()
        {
            if (_ppt == null) return;

            PowerPoint.Slide targetSlide = ResolveActiveSlide();
            if (targetSlide == null) return;

            try
            {
                const float boxWidth = 280f;
                const float boxHeight = 60f;

                float slideWidth = 960f;
                float slideHeight = 540f;
                try
                {
                    var pres = targetSlide.Parent as PowerPoint.Presentation;
                    if (pres?.PageSetup != null)
                    {
                        slideWidth = pres.PageSetup.SlideWidth;
                        slideHeight = pres.PageSetup.SlideHeight;
                    }
                }
                catch { /* ignore */ }

                float left = (slideWidth - boxWidth) / 2f;
                float top = (slideHeight - boxHeight) / 2f;

                var shape = targetSlide.Shapes.AddTextbox(
                    Orientation: Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    Left: left,
                    Top: top,
                    Width: boxWidth,
                    Height: boxHeight);

                var text = shape.TextFrame.TextRange;
                text.Font.Name = "Garamond";
                text.Font.Size = 11;
                text.Text = ""; // keep empty

                // Put caret inside the new textbox
                text.Select();
            }
            catch
            {
                // Swallow errors to avoid crashing PowerPoint.
            }
        }

        private PowerPoint.Slide ResolveActiveSlide()
        {
            try
            {
                PowerPoint.SlideShowView slideShowView = null;
                if (_ppt?.SlideShowWindows?.Count > 0)
                {
                    slideShowView = _ppt.SlideShowWindows[1].View;
                }

                if (slideShowView != null && slideShowView.Slide is PowerPoint.Slide liveSlide)
                {
                    return liveSlide;
                }

                var window = _ppt?.ActiveWindow;
                if (window?.View?.Slide is PowerPoint.Slide windowSlide)
                {
                    return windowSlide;
                }
            }
            catch
            {
                // ignore resolution errors
            }

            return null;
        }

        // ---------------------------------------------------------------------
        // Hotkey plumbing (Ctrl+Shift+T -> InsertStyledTextbox) via window subclassing
        // ---------------------------------------------------------------------

        private void InstallHotkey()
        {
            if (_ppt == null) return;
            if (_subclassHandle != IntPtr.Zero) return;

            _subclassHandle = new IntPtr(_ppt.HWND);
            _wndProcDelegate = WindowProc;

            _originalWndProc = SetWindowLongPtr(_subclassHandle, GWLP_WNDPROC, _wndProcDelegate);

            // Register hotkey Ctrl+Shift+T on the PowerPoint main window
            RegisterHotKey(_subclassHandle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, VK_T);
            // Register hotkey Ctrl+Shift+B to insert styled shape
            RegisterHotKey(_subclassHandle, HOTKEY_SHAPE_ID, MOD_CONTROL | MOD_SHIFT, VK_B);
        }

        private void UninstallHotkey()
        {
            if (_subclassHandle == IntPtr.Zero) return;

            try
            {
                UnregisterHotKey(_subclassHandle, HOTKEY_ID);
                UnregisterHotKey(_subclassHandle, HOTKEY_SHAPE_ID);
            }
            catch
            {
                // ignore
            }

            try
            {
                if (_originalWndProc != IntPtr.Zero)
                {
                    SetWindowLongPtr(_subclassHandle, GWLP_WNDPROC, _originalWndProc);
                }
            }
            catch
            {
                // ignore
            }

            _originalWndProc = IntPtr.Zero;
            _subclassHandle = IntPtr.Zero;
            _wndProcDelegate = null;
        }

        private IntPtr WindowProc(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam)
        {
            if (msg == WM_HOTKEY && wParam.ToInt32() == HOTKEY_ID)
            {
                InsertStyledTextbox();
                return IntPtr.Zero;
            }
            if (msg == WM_HOTKEY && wParam.ToInt32() == HOTKEY_SHAPE_ID)
            {
                InsertCenteredShape();
                return IntPtr.Zero;
            }

            // Listen for Shift+; (:) keydown on the PPT window (without global hotkey)
            const int WM_KEYDOWN = 0x0100;
            if (msg == WM_KEYDOWN && wParam.ToInt32() == VK_OEM_1)
            {
                bool shift = (GetKeyState(VK_SHIFT_KEY) & 0x8000) != 0;
                bool ctrl = (GetKeyState(VK_CONTROL_KEY) & 0x8000) != 0;
                bool alt = (GetKeyState(VK_MENU_KEY) & 0x8000) != 0;

                if (shift && !ctrl && !alt)
                {
                    ShowCommandLineDialog();
                    // Do not consume the key so ':' still types if needed
                }
            }

            return CallWindowProc(_originalWndProc, hWnd, msg, wParam, lParam);
        }

        private bool ShowCommandLineDialog()
        {
            if (ShouldSuppressCmdLine()) return false;

            var slide = ResolveActiveSlide();
            if (slide == null) return false;

            TearDownCmdForm();

            _activeCmdForm = new QuickCommandForm(slide);
            _activeCmdForm.FormClosed += (_, __) => { TearDownCmdForm(); };
            _activeCmdForm.Show();
            return true;
        }

        private bool ShouldSuppressCmdLine()
        {
            try
            {
                var win = _ppt?.ActiveWindow;
                var sel = win?.Selection;
                if (sel == null) return false;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private void TearDownCmdForm()
        {
            try
            {
                if (_activeCmdForm != null)
                {
                    if (!_activeCmdForm.IsDisposed)
                    {
                        _activeCmdForm.Close();
                        _activeCmdForm.Dispose();
                    }
                }
            }
            catch
            {
                // ignore teardown failures
            }
            finally
            {
                _activeCmdForm = null;
            }
        }

        /// <summary>
        /// Inserts a centered rectangle with specified styling and no border.
        /// </summary>
        private void InsertCenteredShape()
        {
            if (_ppt == null) return;
            var slide = ResolveActiveSlide();
            if (slide == null) return;

            try
            {
                // Dimensions in points (4.7" x 0.26")
                const float shapeWidth = 4.7f * 72f;
                const float shapeHeight = 0.26f * 72f;

                float slideWidth = 960f;
                float slideHeight = 540f;
                try
                {
                    var pres = slide.Parent as PowerPoint.Presentation;
                    if (pres?.PageSetup != null)
                    {
                        slideWidth = pres.PageSetup.SlideWidth;
                        slideHeight = pres.PageSetup.SlideHeight;
                    }
                }
                catch
                {
                    // ignore fallback defaults
                }

                float left = (slideWidth - shapeWidth) / 2f;
                float top = (slideHeight - shapeHeight) / 2f;

                var shape = slide.Shapes.AddShape(
                    Type: Office.MsoAutoShapeType.msoShapeRectangle,
                    Left: left,
                    Top: top,
                    Width: shapeWidth,
                    Height: shapeHeight);

                shape.Line.Visible = Office.MsoTriState.msoFalse;
                shape.Fill.ForeColor.RGB = ColorTranslator.ToWin32(Color.FromArgb(0, 32, 96));

                var text = shape.TextFrame.TextRange;
                text.Text = "";
                text.Font.Name = "Garamond";
                text.Font.Size = 12;
                text.Font.Color.RGB = ColorTranslator.ToWin32(Color.White);
                text.Font.Bold = Office.MsoTriState.msoTrue;

                // Select the inserted shape
                try
                {
                    shape.Select();
                }
                catch
                {
                    // ignore selection failures
                }
            }
            catch
            {
                // swallow errors to avoid crashing PowerPoint
            }
        }

        private delegate IntPtr WndProcDelegate(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int GWLP_WNDPROC = -4;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, WndProcDelegate dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll")]
        private static extern IntPtr CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);
    }
}
