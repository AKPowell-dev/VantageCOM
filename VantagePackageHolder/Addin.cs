using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace VantagePackageHolder
{
    [ComVisible(true)]
    [Guid("F5DA47BA-19D6-46CD-ACB7-BC918636925E")]
    [ProgId("VantagePackageHolder.Addin")]
    public sealed class Addin : IDTExtensibility2
    {
        static Addin()
        {
            try
            {
                // Trigger any binding errors early and surface them.
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString(), "VantagePackageHolder.Addin static ctor");
            }
        }

        private Application _excel;
        private VantageEngine _engine;
        private AppEvents_Event _appEvents;
        private IntPtr _subclassHandle = IntPtr.Zero;
        private IntPtr _originalWndProc = IntPtr.Zero;
        private WndProcDelegate _wndProcDelegate;
        private static readonly bool EnableHotkeyHook = true;

        private const int WM_KEYDOWN = 0x0100;
        private const int VK_OEM_4 = 0xDB; // [ key on US keyboards
        private const int VK_OEM_6 = 0xDD; // ] key on US keyboards
        private const int VK_SHIFT_KEY = 0x10;
        private const int VK_CONTROL_KEY = 0x11;
        private const int VK_MENU_KEY = 0x12;

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                _excel = (Application)application;
                _engine = new VantageEngine(_excel);

                if (addInInst is Office.COMAddIn comAddIn)
                {
                    comAddIn.Object = _engine;
                }

                try
                {
                    _appEvents = (AppEvents_Event)_excel;
                    _appEvents.WindowActivate += OnWindowActivate;
                }
                catch
                {
                    // ignore event hookup issues
                }

                if (EnableHotkeyHook)
                {
                    InstallHotkey();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString(), "VantagePackageHolder.OnConnection");
                // Optionally rethrow to preserve failure state
                // throw;
            }
        }

        private void OnWindowActivate(Workbook wb, Window wn)
        {
            try { _engine?.ResetCycleState(); } catch { }
            if (EnableHotkeyHook)
            {
                try { ReinstallHook(wn); } catch { }
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                if (_appEvents != null)
                {
                    _appEvents.WindowActivate -= OnWindowActivate;
                    _appEvents = null;
                }
            }
            catch
            {
                // ignore
            }

            if (EnableHotkeyHook)
            {
                UninstallHotkey();
            }
            _engine?.Dispose();
            _engine = null;
            _excel = null;
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        private void InstallHotkey()
        {
            if (_excel == null) return;
            try { ReinstallHook(_excel.ActiveWindow); } catch { }
        }

        private void ReinstallHook(Window win)
        {
            UninstallHotkey();
            if (win == null) return;
            IntPtr hwnd;
            try { hwnd = new IntPtr(win.Hwnd); } catch { return; }
            if (hwnd == IntPtr.Zero) return;

            _subclassHandle = hwnd;
            _wndProcDelegate = WindowProc;
            _originalWndProc = SetWindowLongPtr(hwnd, GWLP_WNDPROC, _wndProcDelegate);
            if (_originalWndProc == IntPtr.Zero)
            {
                _subclassHandle = IntPtr.Zero;
                _wndProcDelegate = null;
            }
        }

        private void UninstallHotkey()
        {
            if (_subclassHandle == IntPtr.Zero) return;
            try
            {
                if (_originalWndProc != IntPtr.Zero)
                {
                    SetWindowLongPtr(_subclassHandle, GWLP_WNDPROC, _originalWndProc);
                }
            }
            catch { }
            _originalWndProc = IntPtr.Zero;
            _subclassHandle = IntPtr.Zero;
            _wndProcDelegate = null;
        }

        private IntPtr WindowProc(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam)
        {
            if (msg == WM_KEYDOWN && (wParam.ToInt32() == VK_OEM_4 || wParam.ToInt32() == VK_OEM_6))
            {
                bool shift = (GetKeyState(VK_SHIFT_KEY) & 0x8000) != 0;
                bool ctrl = (GetKeyState(VK_CONTROL_KEY) & 0x8000) != 0;
                bool alt = (GetKeyState(VK_MENU_KEY) & 0x8000) != 0;

                if (ctrl && shift && !alt)
                {
                    if (wParam.ToInt32() == VK_OEM_4)
                        TryShowTracePrecedents();
                    else
                        TryShowTraceDependents();
                    return IntPtr.Zero;
                }
            }

            return CallWindowProc(_originalWndProc, hWnd, msg, wParam, lParam);
        }

        private void TryShowTracePrecedents()
        {
            try
            {
                _engine?.TracePrecedentsDialog();
            }
            catch
            {
                // ignore hotkey failures
            }
        }

        private void TryShowTraceDependents()
        {
            try
            {
                _engine?.TraceDependentsDialog();
            }
            catch
            {
                // ignore hotkey failures
            }
        }

        private delegate IntPtr WndProcDelegate(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int GWLP_WNDPROC = -4;

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

