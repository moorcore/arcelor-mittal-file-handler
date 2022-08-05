using IronXL;
using System;
using System.Linq;

namespace ArcelorFileHandler.Services
{
    public class ExcelApi
    {
        public string _startingCellLiteral { get; set; } = "F";
        public string _endingCellLiteral { get; set; } = "F";
        public string _startingCellNumber { get; set; }
        public string _endingCellNumber { get; set; }

        public void GetInvoiceNumbers(string path, string startingCell = "F11", string endingCell = "F1000")
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
        }
    }
}
