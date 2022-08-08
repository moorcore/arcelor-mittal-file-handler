using ArcelorFileHandler.Services;
using System;
using System.IO;
using System.Linq;

namespace ArcelorFileHandler
{
    class Program
    {
        static void Main(string[] args)
        {
            string scanFilesDir = "D:\\Coding\\Arcelor Mittal\\Сентябрь";

            string xlFilesDir = "D:\\Coding\\Arcelor Mittal\\Приложения\\Сентябрь\\";

            FileManager fileManager = new FileManager();
            fileManager.GetInvoiceNumbersFromPath(scanFilesDir);

            ExcelApi excelApi = new ExcelApi();

            foreach (string file in Directory.GetFiles(xlFilesDir))
            {
                excelApi.GetInvoiceNumbers(xlFilesDir, file.Substring(xlFilesDir.Length));
            }

            Console.WriteLine(Enumerable.SequenceEqual
                (fileManager.fileInvoiceNumbersList, excelApi.xlInvoiceNumbersList));
        }
    }
}
