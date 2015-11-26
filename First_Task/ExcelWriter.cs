using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace First_Task
{
    class ExcelWriter
    {
        //private Dictionary<Guid, string> _source;
        private readonly string _fileName;

        public ExcelWriter(string Path)
        {
            _fileName=Path;
        }

        public bool Write(Dictionary<Guid,string> source)
        {
            Excel.Application myExcelApp;
            Excel.Workbook myExcelWorkbook;
            Excel.Worksheet myExcelWorksheet;
            myExcelApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            myExcelWorkbook = myExcelApp.Workbooks.Add(misValue);
            myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.Worksheets.Item[1]; 
                        
            myExcelWorksheet.get_Range("A1", misValue).Formula = "UPC code";
            myExcelWorksheet.get_Range("B1", misValue).Formula = "Link";
            int row = 2;
            foreach(KeyValuePair<Guid,string> pair in source)
            {               
                myExcelWorksheet.Cells[row,1] = pair.Value;
                myExcelWorksheet.Cells[row,2]= @"https://evbyminsd38fe/Assets/Records/~"+pair.Key;
                row++;
            }

            myExcelWorkbook.SaveAs(_fileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                Excel.XlSaveAsAccessMode.xlExclusive,
                misValue, misValue, misValue, misValue, misValue);

            myExcelWorkbook.Close(true, misValue, misValue);
            myExcelApp.Quit();
            return true;
        }
    }
}
