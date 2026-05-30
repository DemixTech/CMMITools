using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BASE.Data
{
    public class OrganizationalUnit
    {
        public string Name {get;set;} // P1PA-SS B31

        public int MaturityLevel { get; set; } // P1PA-SS- B75
        //public void LoadData(Workbook aWkb)
        //{
        //    Worksheet aWks = aWkb.Worksheets["P1PA-SS"];
        //    this.Name = aWks.Cells["B31"].Value;
        //    string maturityLevelStr = aWks.Cells["B75"].Value;
        //    this.MaturityLevel = int.Parse(maturityLevelStr.Substring( startIndex:maturityLevelStr.Length - 1, 1));


        //    //this.Address = $"{aWks.Cells["B18"]}, {aWks.Cells["B19"]}, {aWks.Cells["B20"]}, {aWks.Cells["B21"]}, {aWks.Cells["B22"]}, {aWks.Cells["B23"]}";
        //}
    }
}
