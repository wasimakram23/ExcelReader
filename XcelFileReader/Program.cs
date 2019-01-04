using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel; 


namespace XcelFileReader
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;

            string str = "";
            string path = @"D:\Development\C#\Workplace\Demo.xlsx";
            int row, column, rwCount, coCount;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlRange = xlWorkSheet.UsedRange;
            row = xlRange.Rows.Count;
            column = xlRange.Columns.Count;
            for (rwCount = 1; rwCount <= row; rwCount++)
            {
                for (coCount = 1; coCount <= column; coCount++)
                {
                    object value = (xlRange.Cells[rwCount, coCount] as Excel.Range).Value;
                    Console.Write("{0}\t", value.ToString());
                }
                Console.WriteLine();
            }

        }
    }
}
