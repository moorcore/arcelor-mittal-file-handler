using IronXL;
using System;
using System.Linq;

namespace ArcelorFileHandler.Services
{
    public class ExcelApi
    {
        public void GetInvoiceNumbers(string path)
        {
            int invoiceNumbersCount = 0;

            WorkBook workBook = WorkBook.Load(path);

            foreach (var worsheet in workBook.WorkSheets)
            {
                var array = worsheet["F11:F34"].ToArray();

                for (int i = 0; i < array.Length; i++)
                {
                    var cellValue = array[i];
                    Console.WriteLine(cellValue.ToString().Substring(1));
                    invoiceNumbersCount++;
                }
            }

            /*            WorkSheet ws = workBook.GetWorkSheet("П 42_2021 к Акту №549");

                        var array = ws["F11:F34"].ToArray();

                        for (int i = 0; i < array.Length; i++)
                        {
                            var cellValue = array[i];
                            Console.WriteLine(cellValue.ToString().Substring(1));
                            invoiceNumbersCount++;
                        }*/

            Console.WriteLine(invoiceNumbersCount); // was 24

            Console.Read();
        }
    }
}
