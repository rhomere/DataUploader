using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Uploader.Models
{
    public class Parcel
    {
        public string MunicipalNumber { get; set; }
        public string Owner { get; set; }
        public string Owner2 { get; set; }
        public string MailingAddressLine1 { get; set; }
        public string MailingAddressLine2 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Zip { get; set; }
        public string SiteAddress { get; set; }
        public string StreetNumber { get; set; }
        public string StreetPrefix { get; set; }
        public string StreetName { get; set; }
        public string StreetNumberSuffix { get; set; }
        public string StreetSuffix { get; set; }
        public string CondoUnit { get; set; }
        public string SiteCity { get; set; }
        public string SiteZip { get; set; }
    }
}
