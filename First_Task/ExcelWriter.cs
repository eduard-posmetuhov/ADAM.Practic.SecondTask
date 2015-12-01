using Adam.Core.Records;
using System;
using System.Collections.Generic;
using Adam.Core.Fields;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace First_Task
{
    class ExcelWriter
    {
        private readonly string _fileName;

        public ExcelWriter(string Path)
        {
            _fileName=Path;
        }

        public bool Write(Dictionary<Guid,string> source, Adam.Core.Application app)
        {
            Record r = new Record(app);
            Excel.Application EApp;
            Excel.Workbook EWorkbook;
            Excel.Worksheet EWorksheet;
            Excel.Range Rng;
            EApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            EWorkbook = EApp.Workbooks.Add(misValue);
            EWorksheet = (Excel.Worksheet)EWorkbook.Worksheets.Item[1];                         
            EWorksheet.get_Range("A1", misValue).Formula = "UPC code";
            EWorksheet.get_Range("B1", misValue).Formula = "Link";            
            Rng = EWorksheet.get_Range("A2", misValue).get_Resize(source.Count,misValue);
            Rng.NumberFormat = "00000000000000";
            int row = 2;
            foreach(KeyValuePair<Guid,string> pair in source)
            {              
                EWorksheet.Cells[row,1] = pair.Value;                
                r.Load(pair.Key);
                Rng = EWorksheet.get_Range("B"+row, misValue);
                EWorksheet.Hyperlinks.Add(Rng, r.Fields.GetField<TextField>("Content Url").Value);               
                //myExcelWorksheet.Cells[row, 2] = r.Fields.GetField<TextField>("Content Url").Value;                                
                row++;
            }
            ((Excel.Range)EWorksheet.Cells[2, 1]).EntireColumn.AutoFit();
            ((Excel.Range)EWorksheet.Cells[2, 2]).EntireColumn.AutoFit();
            EWorkbook.SaveAs(_fileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                Excel.XlSaveAsAccessMode.xlExclusive,
                misValue, misValue, misValue, misValue, misValue);

            EWorkbook.Close(true, misValue, misValue);
            EApp.Quit();
            return true;
        }
    }
}
