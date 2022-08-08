using ArcelorFileHandler.Services;
using System;
using System.Linq;

namespace ArcelorFileHandler
{
    class Program
    {
        static void Main(string[] args)
        {
            FileManager fileManager = new FileManager();
            // fileManager.DirSearch("D:\\Coding\\Arcelor Mittal\\Сентябрь");

            string filePath =
                "D:\\Coding\\Arcelor Mittal\\Приложения\\Август\\";

            string fileName = "1_Приложение 1 с 01 по 05 августа 2021 г.xlsx";

            ExcelApi excelApi = new ExcelApi();
            excelApi.GetInvoiceNumbers(filePath, fileName);

            /*Console.WriteLine(Enumerable.SequenceEqual
                (fileManager.fileInvoiceNumbersList, excelApi.xlInvoiceNumbersList));*/
        }
    }
}
