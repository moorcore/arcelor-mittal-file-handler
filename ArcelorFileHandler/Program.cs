using ArcelorFileHandler.Services;
using System;

namespace ArcelorFileHandler
{
    class Program
    {
        static void Main(string[] args)
        {
            FileManager fileManager = new FileManager();
            // fileManager.DirSearch("D:\\Coding\\Arcelor Mittal\\Сентябрь");

            // var invoiceNumbersArray = fileManager.invoiceNumbersList.ToArray();

            string filePath =
                "D:\\Coding\\Arcelor Mittal\\Приложения\\Август\\1_Приложение 1 с 01 по 05 августа 2021 г.xlsx";

            ExcelApi excelApi = new ExcelApi();
            excelApi.GetInvoiceNumbers(filePath);
        }
    }
}
