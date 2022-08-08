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
            FileManager fileManager = new FileManager();
            // fileManager.GetInvoiceNumbersFromPath("D:\\Coding\\Arcelor Mittal\\Сентябрь");

            // TODO: check all files from folder

            string directory =
                "D:\\Coding\\Arcelor Mittal\\Приложения\\Август\\";

            string fileName = "1_Приложение 1 с 01 по 05 августа 2021 г.xlsx";

            ExcelApi excelApi = new ExcelApi();
            // excelApi.GetInvoiceNumbers(directory, fileName);

            foreach (string file in Directory.GetFiles(directory))
            {
                Console.WriteLine(file.Substring(directory.Length));
            }

            /*Console.WriteLine(Enumerable.SequenceEqual
                (fileManager.fileInvoiceNumbersList, excelApi.xlInvoiceNumbersList));*/
        }
    }
}
