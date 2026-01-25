using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace VantageInstaller
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                var scriptPath = Path.Combine(baseDir, "Install-Vantage.ps1");
                if (!ValidatePayload(baseDir, scriptPath))
                {
                    return 1;
                }

                var argList = args ?? Array.Empty<string>();
                var extraArgs = string.Join(" ", argList.Select(EscapeArgument));
                var logPath = Path.Combine(Path.GetTempPath(), "VantageInstall.log");
                var statusPath = Path.Combine(Path.GetTempPath(), "VantageInstall.status.txt");
                File.WriteAllText(statusPath, "Starting installer...");
                var psArgs = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -LogPath \"" + logPath + "\" -StatusPath \"" + statusPath + "\" " + extraArgs;

                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = psArgs,
                    UseShellExecute = true,
                    Verb = "runas",
                    WorkingDirectory = baseDir,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var progressForm = CreateProgressForm())
                {
                    progressForm.Show();
                    UpdateStatusLabel(progressForm.StatusLabel, statusPath);

                    using (var proc = Process.Start(psi))
                    {
                        if (proc == null)
                        {
                            ShowError("Failed to start installer process.");
                            return 1;
                        }

                        while (!proc.HasExited)
                        {
                            UpdateStatusLabel(progressForm.StatusLabel, statusPath);
                            Application.DoEvents();
                            Thread.Sleep(200);
                        }

                        UpdateStatusLabel(progressForm.StatusLabel, statusPath);
                        if (proc.ExitCode != 0)
                        {
                            ShowError("Vantage install failed. See log at:\n" + logPath);
                            return proc.ExitCode;
                        }
                    }
                }

                ShowInfo("Vantage install complete.");
                return 0;
            }
            catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
                ShowError("Install canceled.");
                return 1;
            }
            catch (Exception ex)
            {
                ShowError("Vantage install failed:\n" + ex.Message);
                return 1;
            }
        }

        private static bool ValidatePayload(string baseDir, string scriptPath)
        {
            var missing = new[]
            {
                new { Name = "Install-Vantage.ps1", Ok = File.Exists(scriptPath) },
                new { Name = "VantagePackageHolder.dll", Ok = File.Exists(Path.Combine(baseDir, "VantagePackageHolder.dll")) },
                new { Name = "Extensibility.dll", Ok = File.Exists(Path.Combine(baseDir, "Extensibility.dll")) },
                new { Name = "Vantage.xlam", Ok = File.Exists(Path.Combine(baseDir, "Vantage.xlam")) },
                new { Name = "Resources", Ok = Directory.Exists(Path.Combine(baseDir, "Resources")) }
            }
            .Where(item => !item.Ok)
            .Select(item => item.Name)
            .ToArray();

            if (missing.Length == 0)
            {
                return true;
            }

            var message = "Installer is missing required files next to VantageInstaller.exe:\n"
                          + string.Join("\n", missing)
                          + "\n\nDownload and extract the full VantageInstaller.zip, then run VantageInstaller.exe from that folder.";
            ShowError(message);
            return false;
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(message, "Vantage Installer", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void ShowInfo(string message)
        {
            MessageBox.Show(message, "Vantage Installer", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static ProgressForm CreateProgressForm()
        {
            var form = new ProgressForm();
            return form;
        }

        private static void UpdateStatusLabel(Label label, string statusPath)
        {
            try
            {
                if (!File.Exists(statusPath))
                {
                    return;
                }

                var text = File.ReadAllText(statusPath).Trim();
                if (!string.IsNullOrEmpty(text) && label.Text != text)
                {
                    label.Text = text;
                }
            }
            catch
            {
                // ignore status read failures
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

        private sealed class ProgressForm : Form
        {
            public Label StatusLabel { get; }

            public ProgressForm()
            {
                Text = "Vantage Installer";
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterScreen;
                ControlBox = false;
                TopMost = true;
                Width = 420;
                Height = 140;

                StatusLabel = new Label
                {
                    Dock = DockStyle.Fill,
                    TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                    Padding = new Padding(12, 10, 12, 10),
                    Text = "Starting installer..."
                };

                var progress = new ProgressBar
                {
                    Dock = DockStyle.Bottom,
                    Height = 18,
                    Style = ProgressBarStyle.Marquee
                };

                Controls.Add(StatusLabel);
                Controls.Add(progress);
            }
        }
    }
}
