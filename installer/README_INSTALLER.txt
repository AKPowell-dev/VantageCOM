Vantage Installer

1) Run VantageInstaller.exe as Administrator.
2) If Office is 32-bit, run: VantageInstaller.exe -Force32
   If Office is 64-bit, no flag is needed (or use -Force64).
3) Restart Excel after install.

This installer registers the COM add-in with regasm, copies Vantage.xlam
to %APPDATA%\Microsoft\AddIns, and sets Excel's OPEN registry entry.
