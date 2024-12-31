using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service
{

    public class LibreOfficeService
    {
        private readonly string _libreOfficePath;

        public LibreOfficeService(string libreOfficePath)
        {
            _libreOfficePath = libreOfficePath; // LibreOffice 的執行路徑
        }

        public string ConvertExcelToPdf(string inputFilePath, string outputDirectory)
        {
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException($"File not found: {inputFilePath}");

            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            string outputFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(inputFilePath) + ".pdf");

            // 構建 LibreOffice 命令
            string arguments = $"--headless --convert-to pdf \"{inputFilePath}\" --outdir \"{outputDirectory}\"";

            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = Path.Combine(_libreOfficePath, "soffice.exe"), // LibreOffice 的 soffice 路徑
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            process.WaitForExit();

            // 檢查是否轉換成功
            if (process.ExitCode != 0)
            {
                string error = process.StandardError.ReadToEnd();
                throw new Exception($"LibreOffice conversion failed: {error}");
            }

            return outputFilePath;
        }


        public string ConvertExcelToPdf(string inputFilePath)
        {
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException($"File not found: {inputFilePath}");


            string outputFilePath = inputFilePath.Replace(".xlsx",".pdf");

            // 構建 LibreOffice 命令
            string arguments = $"--headless --convert-to pdf \"{inputFilePath}\" ";

            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = Path.Combine(_libreOfficePath, "soffice.exe"), // LibreOffice 的 soffice 路徑
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            process.WaitForExit();

            // 檢查是否轉換成功
            if (process.ExitCode != 0)
            {
                string error = process.StandardError.ReadToEnd();
                throw new Exception($"LibreOffice conversion failed: {error}");
            }

            return outputFilePath;
        }
    }
}
