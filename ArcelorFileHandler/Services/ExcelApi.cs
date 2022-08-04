using IronXL;
using System;
using System.Linq;

namespace ArcelorFileHandler.Services
{
    public class ExcelApi
    {
        public void GetInvoiceNumbers(string path)
        {
            // TODO: get the number of cells by counting the values from the column with numbers

            int invoiceNumbersCount = 0;

            WorkBook workBook = WorkBook.Load(path);

            foreach (var worksheet in workBook.WorkSheets)
            {
                var array = worksheet["F11:F1000"].ToArray();
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
                        invoiceNumbersCount++;
                    }
                }
            }

            Console.WriteLine();

            Console.WriteLine(invoiceNumbersCount); // 809

            Console.Read();
        }
    }
}
