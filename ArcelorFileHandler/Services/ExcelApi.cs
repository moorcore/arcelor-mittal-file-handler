using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ArcelorFileHandler.Services
{
    public class ExcelApi
    {
        public string _startingCellLiteral { get; set; } = "F";
        public int _startingCellNumber { get; set; } = 11;

        public List<string> xlInvoiceNumbersList = new List<string>();

        public void GetInvoiceNumbers(string path, string fileName)
        {
            string currentCellLiteral = _startingCellLiteral;
            int currentCellNumber = _startingCellNumber;

            string cellValue;

            // int count = 0;

            WorkBook workBook = WorkBook.Load(path + fileName);

            foreach (var worksheet in workBook.WorkSheets)
            {
                var cell = worksheet[$"{currentCellLiteral}{currentCellNumber}"];

/*                Console.WriteLine();
                Console.WriteLine(worksheet.Name);
                Console.WriteLine();*/

                while (!cell.IsEmpty)
                {

                    var firstChar = cell.Value.ToString()[0];

                    if (!char.IsLetter(firstChar))
                    {
                        cellValue = cell.Value.ToString();
                        xlInvoiceNumbersList.Add(cellValue);
                        // count++;
                    }
                    else
                    {
                        cellValue = cell.Value.ToString().Substring(1);
                        xlInvoiceNumbersList.Add(cellValue);
                        // count++;
                    }

                    currentCellNumber++;

                    cell = worksheet[$"{currentCellLiteral}{currentCellNumber}"];
                }

                currentCellNumber = _startingCellNumber;
            }

            // Console.WriteLine(count);

            // Console.WriteLine(xlInvoiceNumbersList.Count);

            // Console.Read();
        }
    }
}
