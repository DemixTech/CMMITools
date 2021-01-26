using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace BASE2.Data
{
    public class Organisation
    {
        public string Name { get; set; } // Organisation's name P1A-SS B16
        public string Address { get; set; } // Organisation's name P1A-SS B18 to B23
    
        public void LoadData(Workbook aWkb)
        {
            Worksheet aWks = (Worksheet)aWkb.Worksheets["P1PA-SS"];
            this.Name = aWks.Cells["B16"].ToString(); // .Value;
            this.Address = $"{aWks.Cells["B18"]}, {aWks.Cells["B19"]}, {aWks.Cells["B20"]}, {aWks.Cells["B21"]}, {aWks.Cells["B22"]}, {aWks.Cells["B23"]}";
        }
    }
}
