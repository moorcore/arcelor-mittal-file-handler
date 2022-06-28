using ArcelorFileHandler.Services;
using System;

namespace ArcelorFileHandler
{
    class Program
    {
        static void Main(string[] args)
        {
            // FileManager.DirSearch("D:\\Coding\\Arcelor Mittal\\Сентябрь");

            string xlFilePath = 
                "D:\\Coding\\Arcelor Mittal\\Приложения\\Август\\1_Приложение 1 с 01 по 05 августа 2021 г";

            ExcelApi excelApi = new ExcelApi(xlFilePath);

            string cellValue = excelApi.GetCellData("Sheet1", 1, 6);
            Console.WriteLine("Cell Value using Column Number: " + cellValue);

            Console.Read();
        }
    }
}
