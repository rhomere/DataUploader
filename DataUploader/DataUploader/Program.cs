using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Uploader;

namespace DataUploader
{
    class Program
    {
        static void Main(string[] args)
        {
            var service = new UploaderService();
            service.LoadWorkbook(@"C:\Users\rhomere\Downloads\sampleTest.xlsx");
            service.ProcessWorksheet(2);
        }
    }
}
