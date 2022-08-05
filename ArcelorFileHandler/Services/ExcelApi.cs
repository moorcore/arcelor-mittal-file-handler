using IronXL;
using System;
using System.Linq;

namespace ArcelorFileHandler.Services
{
    public class ExcelApi
    {
        public string _startingCellLiteral { get; set; } = "F";
        public int _startingCellNumber { get; set; } = 11;

        public void GetInvoiceNumbers(string path)
        {
            string currentCellLiteral = _startingCellLiteral;
            int currentCellNumber = _startingCellNumber;

            int count = 0;

            WorkBook workBook = WorkBook.Load(path);

            foreach (var worksheet in workBook.WorkSheets)
            {
                // TODO: reset the cell number for each worksheet
                var cell = worksheet[$"{currentCellLiteral}{currentCellNumber}"];

                Console.WriteLine();
                Console.WriteLine(worksheet.Name);
                Console.WriteLine();

                while (!cell.IsEmpty)
                {
                    if (cell.Value.ToString().Length >= 7)
                    {
                        var cellValue = cell.Value.ToString();
                        Console.WriteLine(cellValue.Substring(1));
                        count++;
                    }

                    currentCellNumber++;

                    cell = worksheet[$"{currentCellLiteral}{currentCellNumber}"];
                }
            }

            Console.WriteLine();

            Console.WriteLine(count);

            Console.Read();
        }
    }
}
