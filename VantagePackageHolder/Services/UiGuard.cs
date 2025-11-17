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
        private bool _disposed;

        public UiGuard(Excel.Application app, bool hideStatusBar = false)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _prevScreenUpdating = app.ScreenUpdating;
            _prevEnableEvents = app.EnableEvents;
            _hideStatusBar = hideStatusBar;
            _prevStatusBarVisible = app.DisplayStatusBar;

            app.ScreenUpdating = false;
            app.EnableEvents = false;
            if (hideStatusBar)
            {
                app.DisplayStatusBar = false;
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
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
