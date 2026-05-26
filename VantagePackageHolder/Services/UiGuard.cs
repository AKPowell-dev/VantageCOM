using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class UiGuard : IDisposable
    {
        private readonly Excel.Application _app;
        private readonly bool _prevScreenUpdating;
        private readonly bool _prevEnableEvents;
        private readonly bool _hideStatusBar;
        private readonly bool _prevStatusBarVisible;
        private readonly bool _disableAlerts;
        private readonly bool _prevDisplayAlerts;
        private readonly bool _manualCalculation;
        private readonly Excel.XlCalculation _prevCalculation;
        private bool _disposed;

        public UiGuard(Excel.Application app, bool hideStatusBar = false, bool disableAlerts = false, bool manualCalculation = false)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _prevScreenUpdating = app.ScreenUpdating;
            _prevEnableEvents = app.EnableEvents;
            _hideStatusBar = hideStatusBar;
            _prevStatusBarVisible = app.DisplayStatusBar;
            _disableAlerts = disableAlerts;
            _prevDisplayAlerts = app.DisplayAlerts;
            _manualCalculation = manualCalculation;
            _prevCalculation = app.Calculation;

            app.ScreenUpdating = false;
            app.EnableEvents = false;
            if (hideStatusBar)
            {
                app.DisplayStatusBar = false;
            }
            if (disableAlerts)
            {
                app.DisplayAlerts = false;
            }
            if (manualCalculation)
            {
                if (app.Calculation != Excel.XlCalculation.xlCalculationManual)
                {
                    app.Calculation = Excel.XlCalculation.xlCalculationManual;
                }
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            if (_manualCalculation)
            {
                try { _app.Calculation = _prevCalculation; } catch { }
            }
            if (_disableAlerts)
            {
                try { _app.DisplayAlerts = _prevDisplayAlerts; } catch { }
            }
            _app.ScreenUpdating = _prevScreenUpdating;
            _app.EnableEvents = _prevEnableEvents;
            if (_hideStatusBar)
            {
                _app.DisplayStatusBar = _prevStatusBarVisible;
            }

            _disposed = true;
        }
    }
}
