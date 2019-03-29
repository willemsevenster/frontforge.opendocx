using System.Diagnostics;
using System.IO;

namespace DotNetCore.ConsoleApp.Basic
{
    internal class Program
    {
        #region static fields and constants

        private const string FileNameBasicDoc1 = "basic-sample.docx";

        #endregion

        #region implementation

        private static void Main(string[] args)
        {
            // save to file
            using (var fileStream = new FileStream(FileNameBasicDoc1, FileMode.Create))
            {
                BasicSample.Create().Save(fileStream);
                fileStream.Flush();
            }

            // open the file
            var process = new Process
            {
                StartInfo =
                {
                    UseShellExecute = true,
                    FileName = FileNameBasicDoc1
                }
            };

            process.Start();
        }

        #endregion
    }
}