using ArcelorFileHandler.Services;

namespace ArcelorFileHandler
{
    class Program
    {
        // TODO: get the number of cells by couning the values from the column with numbers
        static void Main(string[] args)
        {
            FileManager fileManager = new FileManager();
            // fileManager.DirSearch("D:\\Coding\\Arcelor Mittal\\Сентябрь");

            // var invoiceNumbersArray = fileManager.invoiceNumbersList.ToArray();

            string filePath =
                "D:\\Coding\\Arcelor Mittal\\Приложения\\Август\\1_Приложение 1 с 01 по 05 августа 2021 г.xlsx";

            ExcelApi excelApi = new ExcelApi();
            excelApi.GetInvoiceNumbers(filePath);
        }
    }
}
