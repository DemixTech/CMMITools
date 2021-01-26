using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BASE2.Data;

namespace BASE2.Data
{
    public class Staff
    {
        public string WorkID;
        public string WorkName;
        public string Name; // Participant name
        public string Role; // Participant role
        public List<PracticeArea> PAlist = new List<PracticeArea>();


        // When returning WorkID = null, then this instantiation of Participant object can be discarded
        public void StaffAdd(Worksheet partWks, int row, int headingRow)
        {
            if (partWks == null)
            {
                WorkID = null;
                return;
            }

            string sValue2;

            sValue2 = partWks.Cells[row, 1].ToString();//.Value; // WorkID
            if (string.IsNullOrEmpty(sValue2))
            {
                WorkID = null;
                return;
            }
            else
            {
                WorkID = sValue2.ToLower().Trim(); 
            }

            sValue2 = partWks.Cells[row, 2].ToString();//.Value; // Work name
            if (string.IsNullOrEmpty(sValue2))
            {
                WorkName = "";
            }
            else
            {
                WorkName = sValue2;
            }

            sValue2 = partWks.Cells[row, 3].ToString();//.Value; // Name
            if (string.IsNullOrEmpty(sValue2))
            {
                WorkID = null;
                MessageBox.Show($"Participant name cannot be empty r-{row} c-3");
                return;
            }
            else
            {
                Name = sValue2;
            }

            sValue2 = partWks.Cells[row, 4].ToString();//.Value; // Role
            if (string.IsNullOrEmpty(sValue2))
            {
                Role = "";
            }
            else
            {
                Role = sValue2;
            }


            // *** Populate the PAlist
            string sValue3;
            for (int col = 6; col <= 26; col++)
            {
                // if the col is empty, ignore, else process
                try
                {
                    sValue2 = partWks.Cells[headingRow, col].ToString().ToLower().Trim();  // identify the PA
                    EPAcode thePA = (EPAcode)Enum.Parse(typeof(EPAcode), sValue2, true);
                    PracticeArea aNewPA = new PracticeArea();
                    aNewPA.PAcode = thePA;

                    sValue3 = partWks.Cells[row, col].ToString();
                    if (!string.IsNullOrEmpty(sValue3))
                    {
                        switch (sValue3)
                        {
                            case "x":
                            case "X":
                                aNewPA.SampleType = ESampleType.perform;
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

        public override string ToString()
        {
            return $"{Name} {Role} {WorkID} {PAlist.Count}";
        }
    }
}
