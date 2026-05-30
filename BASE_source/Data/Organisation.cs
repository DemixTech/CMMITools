using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BASE.Data
{
    public class Organisation
    {
        public string Name { get; set; } // Organisation's name P1A-SS B16

        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        public string Country { get; set; }

        public string FullAddress { get { 
            
            return AddressLine1 + ", " + AddressLine2 + ", " + City + ", " + State + ", " + ZipCode + ", " + Country;
            } } // Organisation's name P1A-SS B18 to B23
    
        //public void LoadData(Workbook aWkb)
        //{
        //    Worksheet aWks = aWkb.Worksheets["P1PA-SS"];
        //    this.Name = aWks.Cells["B16"].Value;
        //    this.Address = $"{aWks.Cells["B18"]}, {aWks.Cells["B19"]}, {aWks.Cells["B20"]}, {aWks.Cells["B21"]}, {aWks.Cells["B22"]}, {aWks.Cells["B23"]}";
        //}
    }
}
