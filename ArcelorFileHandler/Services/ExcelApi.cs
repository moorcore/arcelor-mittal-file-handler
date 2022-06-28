using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using xl = Microsoft.Office.Interop.Excel;
using System.Collections;

namespace ArcelorFileHandler.Services
{
    public class ExcelApi
    {
        xl.Application xlApp = null;
        xl.Workbooks workbooks = null;
        xl.Workbook workbook = null;
        Hashtable sheets;
        public string xlFilePath;

        public ExcelApi(string xlFilePath)
        {
            this.xlFilePath = xlFilePath;
        }

        public void OpenExcel()
        {
            xlApp = new xl.Application();
            workbooks = xlApp.Workbooks;
            workbook = workbooks.Open(xlFilePath);
            sheets = new Hashtable();
            int count = 1;

            foreach (xl.Worksheet sheet in workbook.Sheets)
            {
                sheets[count] = sheet.Name;
                count++;
            }
        }

        public void CloseExcel()
        {
            workbook.Close(false, xlFilePath, null);
            Marshal.FinalReleaseComObject(workbook);
            workbook = null;

            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            workbooks = null;

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            xlApp = null;
        }
        public string GetCellData(string sheetName, int colNumber, int rowNumber)
        {
            OpenExcel();

            string value = string.Empty;
            int sheetValue = 0;

            if (sheets.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;

                value = Convert.ToString((range.Cells[rowNumber, colNumber] as xl.Range).Value2);
                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            CloseExcel();
            return value;
        }
    }
}
