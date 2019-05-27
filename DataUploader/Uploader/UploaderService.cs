using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Uploader.Models;
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
                Console.BackgroundColor = ConsoleColor.Green;
                Console.ForegroundColor = ConsoleColor.Black;
                Console.WriteLine("Loading...");
                XlWorkbook = XlApp.Workbooks.Open(path);//@"sandbox_test.xlsx"
                Console.WriteLine("Loading Complete.");
                Console.ResetColor();
            }
            catch(Exception e)
            {
                Console.BackgroundColor = ConsoleColor.Red;
                Console.ForegroundColor = ConsoleColor.Black;
                Console.WriteLine("\nError");
                Console.WriteLine(e.Message);
                Console.WriteLine("Stacktrace");
                Console.WriteLine(e.StackTrace);
                Console.ResetColor();
            }
        }

        public void ProcessWorksheet(int sheetNum)
        {
            try
            {
                Console.BackgroundColor = ConsoleColor.Green;
                Console.ForegroundColor = ConsoleColor.Black;
                Console.WriteLine("\nProcessing...\n");
                Console.ResetColor();
                XlWorksheet = XlWorkbook.Sheets[sheetNum];
                XlRange = XlWorksheet.UsedRange;
                for (int i = 1; i <= XlRange.Rows.Count; i++)
                {

                    var data = new Parcel
                    {
                        MunicipalNumber = XlRange.Cells[i, 1].Value2.ToString(),
                        Owner = XlRange.Cells[i, 2].Value2.ToString(),
                        Owner2 = XlRange.Cells[i, 3].Value2.ToString(),
                        MailingAddressLine1 = XlRange.Cells[i, 4].Value2.ToString(),
                        MailingAddressLine2 = XlRange.Cells[i, 5].Value2.ToString(),
                        City = XlRange.Cells[i, 6].Value2.ToString(),
                        State = XlRange.Cells[i, 7].Value2.ToString(),
                        Zip = XlRange.Cells[i, 8].Value2.ToString(),
                        SiteAddress = XlRange.Cells[i, 9].Value2.ToString(),
                        StreetNumber = XlRange.Cells[i, 10].Value2.ToString(),
                        StreetPrefix = XlRange.Cells[i, 11].Value2.ToString(),
                        StreetName = XlRange.Cells[i, 12].Value2.ToString(),
                        StreetNumberSuffix = XlRange.Cells[i, 13].Value2.ToString(),
                        StreetSuffix = XlRange.Cells[i, 14].Value2.ToString(),
                        CondoUnit = XlRange.Cells[i, 15].Value2.ToString(),
                        SiteCity = XlRange.Cells[i, 16].Value2.ToString(),
                        SiteZip = XlRange.Cells[i, 17].Value2.ToString(),
                    };

                    Console.WriteLine($"{data.MunicipalNumber}, {data.Owner}, {data.MailingAddressLine1}, {data.Owner2}, {data.MailingAddressLine2}, {data.City}, {data.State}, {data.Zip}," +
                        $" {data.SiteAddress}, {data.StreetNumber}, {data.StreetPrefix}, {data.CondoUnit}, {data.SiteCity}, {data.SiteZip}");
                }
                Console.BackgroundColor = ConsoleColor.Green;
                Console.ForegroundColor = ConsoleColor.Black;
                Console.WriteLine("\nProcessing Complete\n");
                Console.ResetColor();
            }
            catch (Exception e)
            {
                Console.BackgroundColor = ConsoleColor.Red;
                Console.ForegroundColor = ConsoleColor.Black;
                Console.WriteLine("\nError");
                Console.WriteLine(e.Message);
                Console.WriteLine("\nStacktrace");
                Console.WriteLine(e.StackTrace);
                Console.ResetColor();
                throw;
            }
            finally
            {
                Cleanup();
            }
        }

        private void Cleanup()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(XlRange);
            Marshal.ReleaseComObject(XlWorksheet);

            //close and release
            XlWorkbook.Close();
            Marshal.ReleaseComObject(XlWorkbook);

            //quit and release
            XlApp.Quit();
            Marshal.ReleaseComObject(XlApp);
        }
    }
}
