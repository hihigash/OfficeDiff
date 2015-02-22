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
            string orignalFileName = args[0];   // 比較元のファイル
            string targetFileName = args[1];    // 比較先のファイル

            string extension = Path.GetExtension(orignalFileName);

            try
            {
                if (String.Compare(extension, ".doc", StringComparison.OrdinalIgnoreCase) == 0 ||
                    String.Compare(extension, ".docx", StringComparison.OrdinalIgnoreCase) == 0 ||
                    String.Compare(extension, ".docm", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    WordComparer.Compare(orignalFileName, targetFileName);
                }
                else if (String.Compare(extension, ".ppt", StringComparison.OrdinalIgnoreCase) == 0 ||
                         String.Compare(extension, ".pptx", StringComparison.OrdinalIgnoreCase) == 0 ||
                         String.Compare(extension, ".pptm", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    PowerPointComparer.Compare(orignalFileName, targetFileName);
                }
                else
                {
                    MessageBox.Show("サポートされていない拡張子です", "エラー");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("アプリケーションの実行中にエラーが発生しました", "エラー");
            }
        }
    }
}
