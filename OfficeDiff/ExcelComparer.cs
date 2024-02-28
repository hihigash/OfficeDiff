using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.Versioning;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OfficeDiff
{
    [SupportedOSPlatform("windows")]
    public class ExcelComparer : IOfficeComparer
    {
        public void Compare(string original, string target)
        {
            string ftype = Registry.ClassesRoot.OpenSubKey(".xls").GetValue(string.Empty) as string;
            string command = Registry.ClassesRoot.OpenSubKey(ftype)
                .OpenSubKey("shell")
                .OpenSubKey("Open")
                .OpenSubKey("command")
                .GetValue(string.Empty) as string;

            var match = Regex.Match(command, @"\""?(.*\.(EXE|exe))\""?");
            if (!match.Success)
            {
                MessageBox.Show("Not found Excel.", "Error");
                return;
            }

            string path = match.Groups[1].Value;
            string appPath = Path.Combine(Path.GetDirectoryName(path), @"DCF\SPREADSHEETCOMPARE.EXE");
            if (!File.Exists(appPath))
            {
                MessageBox.Show("Not found Spreadsheetcompare application.", "Error");
                return;
            }
            string tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.txt");
            File.WriteAllText(tempPath, $"{Path.GetFullPath(original)}{Environment.NewLine}{Path.GetFullPath(target)}");

            ProcessStartInfo info = new ProcessStartInfo(appPath, tempPath);
            Process process = Process.Start(info);
            process.WaitForExit();

            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }
}
