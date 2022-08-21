using IronXL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ArcelorFileHandler.Services
{
    public class ExcelApi
    {
        public string _startingCellLiteral { get; set; } = "F";
        public int _startingCellNumber { get; set; } = 11;

        public List<string> xlInvoiceNumbersList = new List<string>();
        public List<string> xlUniqueInvoiceNumbersList = new List<string>();

        public void GetInvoiceNumbersFromXl(string path, string fileName)
        {
            string currentCellLiteral = _startingCellLiteral;
            int currentCellNumber = _startingCellNumber;
            string cellValue;

            if (!fileName[0].Equals('~'))
            {
                try
                {
                    WorkBook workBook = WorkBook.Load(path + fileName);

                    foreach (var worksheet in workBook.WorkSheets)
                    {
                        if (!worksheet.Name.Equals("Приложение №1"))
                        {
                            var cell = worksheet[$"{currentCellLiteral}{currentCellNumber}"];

                            while (!cell.IsEmpty)
                            {

                                var firstChar = cell.Value.ToString()[0];

                                if (!char.IsLetter(firstChar))
                                {
                                    cellValue = cell.Value.ToString();
                                    xlInvoiceNumbersList.Add(cellValue);
                                }
                                else
                                {
                                    cellValue = cell.Value.ToString().Substring(1);
                                    xlInvoiceNumbersList.Add(cellValue);
                                }

                                currentCellNumber++;

                                cell = worksheet[$"{currentCellLiteral}{currentCellNumber}"];
                            }

                            currentCellNumber = _startingCellNumber;
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }

            xlInvoiceNumbersList.Sort();

            string uniqueNumber = xlInvoiceNumbersList[0];

            for (int i = 1; i < xlInvoiceNumbersList.Count; i++)
            {
                if (xlInvoiceNumbersList[i] != uniqueNumber)
                {
                    uniqueNumber = xlInvoiceNumbersList[i];
                    xlUniqueInvoiceNumbersList.Add(xlInvoiceNumbersList[i]);
                }
            }
        }
    }
}
