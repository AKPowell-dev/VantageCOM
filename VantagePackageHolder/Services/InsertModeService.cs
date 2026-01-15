using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class InsertModeService
    {
        private const int CtrlMask = 512;
        private const int ShiftMask = 1024;
        private const int AltMask = 2048;

        private const int VK_BACK = 8;
        private const int VK_RETURN = 13;
        private const int VK_IME_ON = 25;
        private const int VK_SPACE = 32;
        private const int VK_HOME = 36;
        private const int VK_DELETE = 46;
        private const int VK_F2 = 113;

        private const int VK_LSHIFT = 0xA0;
        private const int VK_RSHIFT = 0xA1;
        private const int VK_LCONTROL = 0xA2;
        private const int VK_RCONTROL = 0xA3;
        private const int VK_LMENU = 0xA4;
        private const int VK_RMENU = 0xA5;

        private const uint KEYEVENTF_EXTENDEDKEY = 0x1;
        private const uint KEYEVENTF_KEYUP = 0x2;

        private readonly Excel.Application _app;

        public InsertModeService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public bool InsertWithIME()
            => RunInsertMode(
                new[] { VK_SPACE, VK_BACK, CtrlMask + VK_HOME, VK_IME_ON },
                new[] { VK_F2, CtrlMask + VK_HOME, VK_IME_ON });

        public bool InsertWithoutIME()
            => RunInsertMode(
                new[] { VK_SPACE, VK_BACK, CtrlMask + VK_HOME },
                new[] { VK_F2, CtrlMask + VK_HOME });

        public bool AppendWithIME()
            => RunInsertMode(
                new[] { VK_SPACE, VK_BACK, VK_IME_ON },
                new[] { VK_F2, VK_IME_ON });

        public bool AppendWithoutIME()
            => RunInsertMode(
                new[] { VK_SPACE, VK_BACK },
                new[] { VK_F2 });

        public bool SubstituteWithIME()
            => RunInsertMode(
                new[] { VK_RETURN, VK_DELETE, VK_IME_ON },
                new[] { VK_BACK, VK_F2, VK_IME_ON });

        public bool SubstituteWithoutIME()
            => RunInsertMode(
                new[] { VK_RETURN, VK_DELETE },
                new[] { VK_BACK, VK_F2 });

        private bool RunInsertMode(int[] shapeSequence, int[] cellSequence)
        {
            bool isRange = IsRangeSelection();

            if (!isRange)
            {
                TryRunMacro("ChangeToShapeInsertMode");
                SendKeySequence(shapeSequence);
                ScheduleDisableIme();
                return false;
            }

            SendKeySequence(cellSequence);
            ScheduleDisableIme();
            return true;
        }

        private bool IsRangeSelection()
        {
            try
            {
                var selection = _app.Selection;
                return selection is Excel.Range;
            }
            catch
            {
                return false;
            }
        }

        private void TryRunMacro(string macroName)
        {
            try
            {
                _app.Run(macroName);
            }
            catch
            {
            }
        }

        private void ScheduleDisableIme()
        {
            try
            {
                _app.OnTime(DateTime.Now.AddMilliseconds(20), "DisableIME");
            }
            catch
            {
            }
        }

        private void SendKeySequence(int[] keys)
        {
            if (keys == null || keys.Length == 0)
            {
                return;
            }

            foreach (int key in keys)
            {
                StrokeSingleKey(key, ignoreKeyUp: false);
            }
        }

        private void StrokeSingleKey(int key, bool ignoreKeyUp)
        {
            bool ctrl = (key & CtrlMask) == CtrlMask;
            bool shift = (key & ShiftMask) == ShiftMask;
            bool alt = (key & AltMask) == AltMask;
            int baseKey = key & 0xFF;

            bool holdAltLeft = false;
            bool holdAltRight = false;
            bool holdCtrlLeft = false;
            bool holdCtrlRight = false;
            bool holdShiftLeft = false;
            bool holdShiftRight = false;

            if (!ignoreKeyUp)
            {
                holdAltRight = IsKeyDown(VK_RMENU) && !alt;
                holdCtrlRight = IsKeyDown(VK_RCONTROL) && !ctrl;
                holdShiftRight = IsKeyDown(VK_RSHIFT) && !shift;

                if (!alt)
                {
                    holdAltLeft = IsKeyDown(VK_LMENU);
                }

                if (!ctrl)
                {
                    holdCtrlLeft = IsKeyDown(VK_LCONTROL);
                }

                if (!shift)
                {
                    holdShiftLeft = IsKeyDown(VK_LSHIFT);
                }

                if (holdAltRight)
                {
                    keybd_event((byte)VK_RMENU, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, UIntPtr.Zero);
                }

                if (holdCtrlRight)
                {
                    keybd_event((byte)VK_RCONTROL, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, UIntPtr.Zero);
                }

                if (holdShiftRight)
                {
                    keybd_event((byte)VK_RSHIFT, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, UIntPtr.Zero);
                }

                if (holdAltLeft)
                {
                    keybd_event((byte)VK_LMENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                }

                if (holdCtrlLeft)
                {
                    keybd_event((byte)VK_LCONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                }

                if (holdShiftLeft)
                {
                    keybd_event((byte)VK_LSHIFT, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                }
            }

            if (alt)
            {
                keybd_event((byte)VK_RMENU, 0, KEYEVENTF_EXTENDEDKEY, UIntPtr.Zero);
            }

            if (ctrl)
            {
                keybd_event((byte)VK_RCONTROL, 0, KEYEVENTF_EXTENDEDKEY, UIntPtr.Zero);
            }

            if (shift)
            {
                keybd_event((byte)VK_RSHIFT, 0, KEYEVENTF_EXTENDEDKEY, UIntPtr.Zero);
            }

            if (baseKey > 0)
            {
                keybd_event((byte)baseKey, 0, 0, UIntPtr.Zero);
                keybd_event((byte)baseKey, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            }

            if (shift && !holdShiftRight)
            {
                keybd_event((byte)VK_RSHIFT, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, UIntPtr.Zero);
            }

            if (ctrl && !holdCtrlRight)
            {
                keybd_event((byte)VK_RCONTROL, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, UIntPtr.Zero);
            }

            if (alt && !holdAltRight)
            {
                keybd_event((byte)VK_RMENU, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, UIntPtr.Zero);
            }

            if (ignoreKeyUp)
            {
                return;
            }

            if (holdShiftLeft)
            {
                keybd_event((byte)VK_LSHIFT, 0, 0, UIntPtr.Zero);
            }

            if (holdCtrlLeft)
            {
                keybd_event((byte)VK_LCONTROL, 0, 0, UIntPtr.Zero);
            }

            if (holdAltLeft)
            {
                keybd_event((byte)VK_LMENU, 0, 0, UIntPtr.Zero);
            }

            if (holdShiftRight && !shift)
            {
                keybd_event((byte)VK_RSHIFT, 0, 0, UIntPtr.Zero);
            }

            if (holdCtrlRight && !ctrl)
            {
                keybd_event((byte)VK_RCONTROL, 0, 0, UIntPtr.Zero);
            }

            if (holdAltRight && !alt)
            {
                keybd_event((byte)VK_RMENU, 0, 0, UIntPtr.Zero);
            }
        }

        private static bool IsKeyDown(int key)
            => (GetKeyState(key) & 0x8000) != 0;

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int vKey);
    }
}
