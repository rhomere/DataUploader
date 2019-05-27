using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Uploader
{
    public class UploaderService
    {
        Excel.Application XlApp = new Excel.Application();
        Excel.Workbook XlWorkbook = null;// new Excel.Workbook();
        Excel._Worksheet XlWorksheet = null;// new Excel.Worksheet();
        Excel.Range XlRange = null;

        public UploaderService()
        {
            
        }

        public void LoadWorkbook(string path)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Loading...");
                XlWorkbook = XlApp.Workbooks.Open(path);//@"sandbox_test.xlsx"
                Console.WriteLine("Loading Complete.");
            }
            catch(Exception e)
            {
                Console.BackgroundColor = ConsoleColor.White;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nError");
                Console.WriteLine(e.Message);
                Console.WriteLine("Stacktrace");
                Console.WriteLine(e.StackTrace);
                Console.ResetColor();
            }
        }

        public void ProcessWorksheet(int sheetNum)
        {
            XlWorksheet = XlWorkbook.Sheets[sheetNum];
            XlRange = XlWorksheet.UsedRange;

            for (int i = 1; i <= XlRange.Rows.Count; i++)
            {
                for (int j = 1; j <= XlRange.Columns.Count; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    if (XlRange.Cells[i, j] != null && XlRange.Cells[i, j].Value2 != null)
                        Console.Write(XlRange.Cells[i, j].Value2.ToString() + "\t");

                    //add useful things here!   
                }
            }
        }
    }
}
