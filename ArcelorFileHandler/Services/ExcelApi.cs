using IronXL;
using System;
using System.Linq;

namespace ArcelorFileHandler.Services
{
    public class ExcelApi
    {
        public string _startingCellLiteral { get; set; } = "F";
        public string _endingCellLiteral { get; set; } = "F";
        public int _startingCellNumber { get; set; } = 11;
        public int _endingCellNumber { get; set; } = 1000;

        public void GetInvoiceNumbers(string path, string startingCell = "F11", string endingCell = "F1000")
        {
            startingCell = $"{_startingCellLiteral}{_startingCellNumber}";
            endingCell = $"{_endingCellLiteral}{_endingCellNumber}";

            string currentCellLiteral = _startingCellLiteral;
            int currentCellNumber = _startingCellNumber;

            WorkBook workBook = WorkBook.Load(path);

            foreach (var worksheet in workBook.WorkSheets)
            {
                var cell = worksheet[$"{currentCellLiteral}{currentCellNumber}"];

                while (!cell.IsEmpty)
                {
                    if (cell.Value.ToString().Length >= 7)
                    {
                        var cellValue = cell.Value.ToString();
                        Console.WriteLine(cellValue.Substring(1));
                    }

                    currentCellNumber++;

                    cell = worksheet[$"{currentCellLiteral}{currentCellNumber}"];
                }
            }

            Console.WriteLine();

            Console.Read();
        }

/*        public void GetInvoiceNumbers(string path, string startingCell = "F11", string endingCell = "F1000")
        {
            startingCell = $"{_startingCellLiteral}{_startingCellNumber}";
            endingCell = $"{_endingCellLiteral}{_endingCellNumber}";

            WorkBook workBook = WorkBook.Load(path);

            foreach (var worksheet in workBook.WorkSheets)
            {
                var array = worksheet[$"{startingCell}:{endingCell}"].ToArray();
                Console.WriteLine(worksheet.RowCount);

                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i].IsEmpty)
                    {
                        continue;
                    }

                    if (array[i].Text.Length >= 7)
                    {
                        var cellValue = array[i];
                        Console.WriteLine(cellValue.ToString().Substring(1));
                    }
                }
            }

            Console.WriteLine();

            Console.Read();
        }*/
    }
}
