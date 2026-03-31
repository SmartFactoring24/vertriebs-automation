using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        try
        {
            var exeDir = AppDomain.CurrentDomain.BaseDirectory;
            var scriptPath = Path.Combine(exeDir, "StartKiAutomation.ps1");

            if (!File.Exists(scriptPath))
            {
                MessageBox.Show(
                    "Die Startdatei wurde nicht gefunden:\n" + scriptPath,
                    "StartKiAutomation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\"",
                WorkingDirectory = exeDir,
                UseShellExecute = true
            };

            Process.Start(startInfo);
        }
        catch (Exception exception)
        {
            MessageBox.Show(
                exception.ToString(),
                "StartKiAutomation",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }
}
