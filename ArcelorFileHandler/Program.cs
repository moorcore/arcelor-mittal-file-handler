using ArcelorFileHandler.Services;
using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ArcelorFileHandler
{
    class Program
    {
        // TODO: get the number of cells by couning the values from the column with numbers
        static void Main(string[] args)
        {
            // FileManager.DirSearch("D:\\Coding\\Arcelor Mittal\\Сентябрь");

            string filePath =
                "D:\\Coding\\Arcelor Mittal\\Приложения\\Август\\1_Приложение 1 с 01 по 05 августа 2021 г.xlsx";

            WorkBook wb = WorkBook.Load(filePath);
            WorkSheet ws = wb.GetWorkSheet("П 42_2021 к Акту №549");

            var array = ws["F11:F34"].ToArray();

            for (int i = 0; i < array.Length; i++)
            {
                var cellValue = array[i];
                Console.WriteLine(cellValue.ToString().Substring(1));
            }

            Console.Read();
        }
    }
}
