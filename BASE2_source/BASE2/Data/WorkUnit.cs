using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BASE2.Data;
using Range = Microsoft.Office.Interop.Excel.Range;


namespace BASE2.Data
{
    public class WorkUnit
    {
        public EWorkType WorkType { get; set; } 
        public string ID;
        public string Name;
        public string Description;
        public string Lifecycle; // group
        public string Stage; // 
        public DateTime StartDate;
        public DateTime EndDate;
        public bool Include; // include in sample, yes/no
        public List<PracticeArea> PAlist = new List<PracticeArea>();

        // When returning WorkType = EWorkTYpe.nothing, then this instantiation of WorkUnit object can be discarded
        public WorkUnit() //EWorkType workType2, Worksheet prjWks, int row, int headingRow)
        {
            


        }
        public void AddWorkType(EWorkType workType2, Worksheet prjWks, int row, int headingRow) 
        {
            if (prjWks == null)
            {
                WorkType = EWorkType.nothing;
                return;
            }

            switch (workType2)
            {
                case EWorkType.project:
                    ProcessProjects(prjWks, row, headingRow);
                    break;
                case EWorkType.support:
                    ProcessSupport(prjWks, row, headingRow);
                    break;
            }
        }

        // ***** HELPER FUNCTIONS
        private void ProcessProjects(Worksheet prjWks, int row, int headingRow)
        {

            //string sValue2;
            try
            {
                Range aRange1 = (Range)prjWks.Cells[row, 1];
                string sValue1 = aRange1.Value.ToString(); // (Range)(prjWks.Cells[row, 1]) .Value.ToString();//.Value; // ID field
                if (string.IsNullOrEmpty(sValue1))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    ID = sValue1.ToLower().Trim();
                    WorkType = EWorkType.project;
                }

                Range aRange2 = (Range)prjWks.Cells[row, 2];
                string sValue2 = aRange2.Value.ToString(); // prjWks.Cells[row, 2].ToString();//.Value; // Name
                if (string.IsNullOrEmpty(sValue2))
                {
                    WorkType = EWorkType.nothing;
                    MessageBox.Show($"Project sheet, name field is empty r{row} c2");
                    return;
                }
                else
                {
                    Name = sValue2;
                }

                Range aRange3 = (Range)prjWks.Cells[row, 3];
                string sValue32 = aRange3.Value.ToString();// prjWks.Cells[row, 3].ToString();//.Value; // Description
                if (string.IsNullOrEmpty(sValue32))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    Description = sValue32;
                }

                Range aRange4 = (Range)prjWks.Cells[row, 4];
                string sValue4 = aRange4.Value.ToString();

                //sValue2 = prjWks.Cells[row, 4].ToString();//.Value; // Lifecycle
                if (string.IsNullOrEmpty(sValue4))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    Lifecycle = sValue4;
                }

                Range aRange5 = (Range)prjWks.Cells[row, 5];
                string sValue5 = aRange5.Value.ToString();
                //sValue2 = prjWks.Cells[row, 5].ToString();//.Value; // Stage
                if (string.IsNullOrEmpty(sValue5))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    Stage = sValue5;
                }

                Range aRange6 = (Range)prjWks.Cells[row, 6];
                string sValue6 = aRange6.Value.ToString();
                //sValue2 = prjWks.Cells[row, 6].ToString();//.Value.ToString(); // Start
                if (string.IsNullOrEmpty(sValue6))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    StartDate = DateTime.Parse(sValue6);
                }

                Range aRange7 = (Range)prjWks.Cells[row, 7];
                string sValue7 = aRange7.Value.ToString();
                //sValue2 = prjWks.Cells[row, 7].ToString();//.Value.ToString(); // End
                if (string.IsNullOrEmpty(sValue7))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    EndDate = DateTime.Parse(sValue7);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not read line:{row} in tab:Projects. Error{ex.Message}");
            }
            // *** Populate the PAlist
            string sValue3;
            for (int col = 10; col <= 27; col++)
            {
                // if the col is empty, ignore, else process
                try
                {
                   // var aCell = prjWks.Cells[headingRow, col];//.Value;
                    Range aCellRnd = (Range)prjWks.Cells[headingRow, col];//
                    string sValue2 = aCellRnd.Value.ToString();
                    if (sValue2 == null) //aCellRnd == null) //aCell == null)
                    {

                    }
                    else
                    {
                        //sValue2 = (string)prjWks.Cells[1, col].Value.ToString().ToLower().Trim();  // identify the PA
                        //sValue2 = aCell.ToString().ToLower().Trim();
                        EPAcode thePA = (EPAcode)Enum.Parse(typeof(EPAcode), sValue2, true);
                        PracticeArea aNewPA = new PracticeArea();
                        aNewPA.PAcode = thePA;

                      //  var aCell3 = prjWks.Cells[row, col];//.Value;
                        Range aCellRange3 = (Range)prjWks.Cells[row, col]; //
                        if (aCellRange3 == null) // aCell3 == null)
                        {
                            var xxxx = 1;
                        }
                        else
                        {

                            //sValue3 = (string)prjWks.Cells[row, col].Value.ToString();
                           // sValue3 = (string)aCell3.ToString();
                            sValue3 = aCellRange3.ToString();
                            if (!string.IsNullOrEmpty(sValue3))
                            {
                                switch (sValue3)
                                {
                                    case "x":
                                    case "X":
                                    case "s":
                                    case "S":
                                        aNewPA.SampleType = ESampleType.sampled;
                                        PAlist.Add(aNewPA);
                                        break;
                                    case "a":
                                    case "A":
                                        aNewPA.SampleType = ESampleType.added;
                                        PAlist.Add(aNewPA);
                                        break;
                                    default:
                                        aNewPA.SampleType = ESampleType.not;
                                        break;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // If exception, then it could not match, ignore
                    MessageBox.Show($"Could not read PA in row:{row} col:{col} in tab:projects. Error {ex.Message}");
                    Debug.WriteLine($"Error: {ex.Message}");
                }



            }
        }
        private void ProcessSupport(Worksheet prjWks, int row, int headingRow)
        {
            //string sValue2;
            string aName = prjWks.Name;
            Range aRange1 = (Range)prjWks.Cells[row, 1];// .Range[row, 1];
            string sValue1 = aRange1?.Value?.ToString();
            //sValue2 = prjWks.Cells[row, 1].ToString();//.Value; // ID
            if (string.IsNullOrEmpty(sValue1)) // string.IsNullOrEmpty(sValue2))
            {
                WorkType = EWorkType.nothing;
                return;
            }
            else
            {
                ID = sValue1.ToLower().Trim(); // sValue2.ToLower().Trim();
                WorkType = EWorkType.support;
            }

            Range aRange2 = (Range)prjWks.Cells[row, 2];
            string sValue2 = aRange2?.Value?.ToString();//  aRange2.Value.ToString();
           // sValue2 = prjWks.Cells[row, 2].ToString();//.Value; // Name
            if (string.IsNullOrEmpty(sValue2))
            {
                WorkType = EWorkType.nothing;
                MessageBox.Show($"Support sheet, name field is empty r{row} c2");
                return;
            }
            else
            {
                Name = sValue2;
            }

            //sValue2 = prjWks.Cells[row, 3].ToString();//.Value; // Description
            Range aRange3 = (Range)prjWks.Cells[row, 3];
            string sValue32 = aRange3?.Value?.ToString();// prjWks.Cells[row, 3].Value.ToString();
            if (string.IsNullOrEmpty(sValue32))
            {
                Description = "";
            }
            else
            {
                Description = sValue32;
            }

            //sValue2 = prjWks.Cells[row, 4].ToString();//.Value; // Lifecycle
            Range aRange4 = (Range)prjWks.Cells[row, 4];
            string sValue4 = aRange4?.Value?.ToString();// prjWks.Cells[row, 4].Value.ToString();
            if (string.IsNullOrEmpty(sValue4))
            {
                Lifecycle = "";
            }
            else
            {
                Lifecycle = sValue4;
            }

            // *** Populate the PAlist
            //string sValuePAstr;
            for (int col = 6; col <= 26; col++)
            {

                // if the col is empty, ignore, else process
                try
                {
                    Range aRangePA1 = (Range)prjWks.Cells[headingRow, col];
                    string sValuePAstr1 = aRangePA1?.Value.ToString().ToLower().Trim() ?? "";

                    //sValuePAstr = prjWks.Cells[headingRow, col].ToString().ToLower().Trim();//.Cells[headingRow, col].ToString().ToLower().Trim();  // identify the PA
                    EPAcode thePA = (EPAcode)Enum.Parse(typeof(EPAcode), sValuePAstr1, true);// sValue2, true);
                    PracticeArea aNewPA = new PracticeArea();
                    aNewPA.PAcode = thePA;

                    Range aRangePA2 = (Range)prjWks.Cells[row, col];
                    string sValuePAstr2 = aRangePA2?.Value?.ToString() ?? "";

                    //sValue3 = prjWks.Cells[row, col].Value.ToString(); ;// prjWks.Cells[row, col].ToString();
                    if (!string.IsNullOrEmpty(sValuePAstr2)) //sValue3))
                    {
                        switch (sValuePAstr2) //sValue3)
                        {
                            case "x":
                            case "X":
                            case "s":
                            case "S":
                                aNewPA.SampleType = ESampleType.sampled;
                                PAlist.Add(aNewPA);
                                break;
                            case "a":
                            case "A":
                                aNewPA.SampleType = ESampleType.added;
                                PAlist.Add(aNewPA);
                                break;
                            default:
                                aNewPA.SampleType = ESampleType.not;
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // If exception, then it could not match, ignore
                    Debug.WriteLine($"Error: {ex.Message}");
                }



            }
        }
    }

    public enum EWorkType
    {
        nothing = 0, // nothing to process, do not add to list
        project = 1,
        support = 2
    }


}
