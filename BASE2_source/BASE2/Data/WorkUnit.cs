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

            string sValue2;
            try
            {
                sValue2 = prjWks.Cells[row, 1].ToString();//.Value; // ID field
                if (string.IsNullOrEmpty(sValue2))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    ID = sValue2.ToLower().Trim();
                    WorkType = EWorkType.project;
                }

                sValue2 = prjWks.Cells[row, 2].ToString();//.Value; // Name
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

                sValue2 = prjWks.Cells[row, 3].ToString();//.Value; // Description
                if (string.IsNullOrEmpty(sValue2))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    Description = sValue2;
                }

                sValue2 = prjWks.Cells[row, 4].ToString();//.Value; // Lifecycle
                if (string.IsNullOrEmpty(sValue2))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    Lifecycle = sValue2;
                }

                sValue2 = prjWks.Cells[row, 5].ToString();//.Value; // Stage
                if (string.IsNullOrEmpty(sValue2))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    Stage = sValue2;
                }

                sValue2 = prjWks.Cells[row, 6].ToString();//.Value.ToString(); // Start
                if (string.IsNullOrEmpty(sValue2))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    StartDate = DateTime.Parse(sValue2);
                }

                sValue2 = prjWks.Cells[row, 7].ToString();//.Value.ToString(); // End
                if (string.IsNullOrEmpty(sValue2))
                {
                    WorkType = EWorkType.nothing;
                    return;
                }
                else
                {
                    EndDate = DateTime.Parse(sValue2);
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
                    var aCell = prjWks.Cells[headingRow, col];//.Value;
                    if (aCell == null)
                    {

                    }
                    else
                    {
                        //sValue2 = (string)prjWks.Cells[1, col].Value.ToString().ToLower().Trim();  // identify the PA
                        sValue2 = aCell.ToString().ToLower().Trim();
                        EPAcode thePA = (EPAcode)Enum.Parse(typeof(EPAcode), sValue2, true);
                        PracticeArea aNewPA = new PracticeArea();
                        aNewPA.PAcode = thePA;

                        var aCell3 = prjWks.Cells[row, col];//.Value;
                        if (aCell3 == null)
                        {
                            var xxxx = 1;
                        }
                        else
                        {

                            //sValue3 = (string)prjWks.Cells[row, col].Value.ToString();
                            sValue3 = (string)aCell3.ToString();
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
            string sValue2;
            sValue2 = prjWks.Cells[row, 1].ToString();//.Value; // ID
            if (string.IsNullOrEmpty(sValue2))
            {
                WorkType = EWorkType.nothing;
                return;
            }
            else
            {
                ID = sValue2.ToLower().Trim();
                WorkType = EWorkType.support;
            }

            sValue2 = prjWks.Cells[row, 2].ToString();//.Value; // Name
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

            sValue2 = prjWks.Cells[row, 3].ToString();//.Value; // Description
            if (string.IsNullOrEmpty(sValue2))
            {
                Description = "";
            }
            else
            {
                Description = sValue2;
            }

            sValue2 = prjWks.Cells[row, 4].ToString();//.Value; // Lifecycle
            if (string.IsNullOrEmpty(sValue2))
            {
                Lifecycle = "";
            }
            else
            {
                Lifecycle = sValue2;
            }

            // *** Populate the PAlist
            string sValue3;
            for (int col = 6; col <= 26; col++)
            {

                // if the col is empty, ignore, else process
                try
                {
                    sValue2 = prjWks.Cells[headingRow, col].ToString().ToLower().Trim();  // identify the PA
                    EPAcode thePA = (EPAcode)Enum.Parse(typeof(EPAcode), sValue2, true);
                    PracticeArea aNewPA = new PracticeArea();
                    aNewPA.PAcode = thePA;

                    sValue3 = prjWks.Cells[row, col].ToString();
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
