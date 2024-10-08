﻿using ArcelorFileHandler.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ArcelorFileHandler
{
    class Program
    {
        static void Main(string[] args)
        {
            string scanFilesDir = "D:\\Coding\\Arcelor Mittal\\Сентябрь";

            string xlFilesDir = "D:\\Coding\\Arcelor Mittal\\Приложения\\Сентябрь\\";

            FileManager fileManager = new FileManager();
            fileManager.GetInvoiceNumbersFromFileName(scanFilesDir);

            ExcelApi excelApi = new ExcelApi();

            foreach (string file in Directory.GetFiles(xlFilesDir))
            {
                excelApi.GetInvoiceNumbersFromXl(xlFilesDir, file.Substring(xlFilesDir.Length));
            }

            Console.WriteLine(fileManager.fileInvoiceNumbersList.Count);
            Console.WriteLine(excelApi.xlUniqueInvoiceNumbersList.Count);

            Console.WriteLine();

            Console.WriteLine(Enumerable.SequenceEqual
                (fileManager.fileInvoiceNumbersList, excelApi.xlUniqueInvoiceNumbersList));
        }
    }
}
