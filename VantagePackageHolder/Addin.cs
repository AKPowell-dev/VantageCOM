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
                    _appEvents.SheetSelectionChange += OnSheetSelectionChange;
                    _appEvents.WindowActivate += OnWindowActivate;
                }
                catch
                {
                    // ignore event hookup issues
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString(), "VantagePackageHolder.OnConnection");
                // Optionally rethrow to preserve failure state
                // throw;
            }
        }

        private void OnSheetSelectionChange(object sh, Range target)
        {
            try { _engine?.ResetCycleState(); } catch { }
        }

        private void OnWindowActivate(Workbook wb, Window wn)
        {
            try { _engine?.ResetCycleState(); } catch { }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                if (_appEvents != null)
                {
                    _appEvents.SheetSelectionChange -= OnSheetSelectionChange;
                    _appEvents.WindowActivate -= OnWindowActivate;
                    _appEvents = null;
                }
            }
            catch
            {
                // ignore
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
    }
}

