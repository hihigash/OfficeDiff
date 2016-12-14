// ===========================================================================
// Program.cs
// ---------------------------------------------------------------------------
// This Sample Code is provided for the purpose of illustration only and is 
// not intended to be used in a production environment.  THIS SAMPLE CODE AND 
// ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
// EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
// WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We 
// grant You a nonexclusive, royalty-free right to use and modify the Sample 
// Code and to reproduce and distribute the object code form of the Sample 
// Code, provided that You agree: (i) to not use Our name, logo, or trademarks
// to market Your software product in which the Sample Code is embedded; (ii) 
// to include a valid copyright notice on Your software product in which the 
// Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend 
// Us and Our suppliers from and against any claims or lawsuits, including 
// attorneys' fees, that arise or result from the use or distribution of the 
// Sample Code.
// ===========================================================================
using System;
using System.IO;
using System.Windows.Forms;

namespace OfficeDiff
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                MessageBox.Show("OfficeDiff.exe <Compare file path> <To file path>","Usage");
                return;
            }

            string orignalFileName = args[0];   // source file
            string targetFileName = args[1];    // target file

            string extension = Path.GetExtension(orignalFileName);

            try
            {
                IOfficeComparer comparer = null;
                if (string.Compare(extension, ".doc", StringComparison.OrdinalIgnoreCase) == 0 ||
                    string.Compare(extension, ".docx", StringComparison.OrdinalIgnoreCase) == 0 ||
                    string.Compare(extension, ".docm", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    comparer = new WordComparer();
                }
                else if (string.Compare(extension, ".ppt", StringComparison.OrdinalIgnoreCase) == 0 ||
                         string.Compare(extension, ".pptx", StringComparison.OrdinalIgnoreCase) == 0 ||
                         string.Compare(extension, ".pptm", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    comparer = new PowerPointComparer();
                }
                else if (string.Compare(extension, ".xls", StringComparison.OrdinalIgnoreCase) == 0 ||
                         string.Compare(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) == 0 ||
                         string.Compare(extension, ".xlsm", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    comparer = new ExcelComparer();
                }
                else
                {
                    MessageBox.Show("Unsupported extension types", "Error");
                }

                comparer?.Compare(orignalFileName, targetFileName);
            }
            catch (Exception)
            {
                MessageBox.Show("An error occurred while executing the application", "Error");
            }
        }
    }
}
