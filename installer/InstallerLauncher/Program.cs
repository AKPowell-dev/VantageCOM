using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace VantageInstaller
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            try
            {
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                var scriptPath = Path.Combine(baseDir, "Install-Vantage.ps1");
                if (!File.Exists(scriptPath))
                {
                    Console.Error.WriteLine("Install-Vantage.ps1 was not found next to the installer.");
                    return 1;
                }

                var argList = args ?? Array.Empty<string>();
                var extraArgs = string.Join(" ", argList.Select(EscapeArgument));
                var psArgs = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" " + extraArgs;

                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = psArgs,
                    UseShellExecute = true,
                    Verb = "runas",
                    WorkingDirectory = baseDir
                };

                using (var proc = Process.Start(psi))
                {
                    proc.WaitForExit();
                    return proc.ExitCode;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                return 1;
            }
        }

        private static string EscapeArgument(string arg)
        {
            if (string.IsNullOrEmpty(arg))
            {
                return "\"\"";
            }

            if (arg.IndexOf(' ') >= 0 || arg.IndexOf('"') >= 0)
            {
                return "\"" + arg.Replace("\"", "\\\"") + "\"";
            }

            return arg;
        }
    }
}
