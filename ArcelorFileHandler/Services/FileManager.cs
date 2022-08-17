using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ArcelorFileHandler.Services
{
    public class FileManager
    {
        public List<string> fileInvoiceNumbersList = new List<string>();

        public void GetInvoiceNumbersFromFileName(string path)
        {
            string invoiceNumber;

            try
            {
                foreach (string dir in Directory.GetDirectories(path))
                {
                    foreach (string file in Directory.GetFiles(dir))
                    {
                        var lastOperatorIndex = file.LastIndexOf(' ');
                        string tempString = file.Substring(lastOperatorIndex);
                        invoiceNumber = tempString.Substring(1, tempString.Length - 5);

                        if (invoiceNumber.Length >= 7)
                        {
                            fileInvoiceNumbersList.Add(invoiceNumber);
                        }
                    }
                    GetInvoiceNumbersFromFileName(dir);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
