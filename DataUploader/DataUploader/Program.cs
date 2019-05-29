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
            var dbService = new DbService();
            service.LoadWorkbook(@"C:\Users\User\Downloads\sample.xlsx");
            service.ProcessWorksheet(1);
            var data = service.GetParcelData();
            foreach (var item in data)
            {
                dbService.Add(item);
            }
        }
    }
}
