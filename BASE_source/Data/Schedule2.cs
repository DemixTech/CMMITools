using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BASE
{
    public class Schedule2
    {
        //public EWorkType WorkType;
        public string WorkID;
        public string PA;
        public string ParticipantName;
        public string Role;
        public string InterviewSession;

        
        public void Schedule2Add(Worksheet schWks, int row)
        {
            if (schWks == null)
            {
                WorkID = null;
                return;
            }

            string sValue2;
            try
            {
                sValue2 = (string)schWks.Cells[row, 1].Value; // WorkID field
                if (string.IsNullOrEmpty(sValue2))
                {
                    WorkID = null;
                    return;
                }
                else
                {
                    WorkID = sValue2.ToLower().Trim();
                }

                sValue2 = (string)schWks.Cells[row, 3].Value; // PA
                if (string.IsNullOrEmpty(sValue2))
                {
                    MessageBox.Show($"Schedule2 sheet, PA field empty r{row} c3");
                    WorkID = null;
                    return;
                }
                else
                {
                    PA = sValue2;
                }

                sValue2 = (string)schWks.Cells[row, 4].Value; // Participant Name
                if (string.IsNullOrEmpty(sValue2))
                {
                    MessageBox.Show($"Schedule2 sheet, PA field empty r{row} c4");
                    WorkID = null;
                    return;
                }
                else
                {
                    ParticipantName = sValue2;
                }

                sValue2 = (string)schWks.Cells[row, 5].Value; // Role
                if (string.IsNullOrEmpty(sValue2))
                {
                    MessageBox.Show($"Schedule2 sheet, PA field empty r{row} c5");
                    WorkID = null;
                    return;
                }
                else
                {
                    Role = sValue2;
                }

                sValue2 = (string)schWks.Cells[row, 8].Value; // PA
                if (string.IsNullOrEmpty(sValue2))
                {
                    MessageBox.Show($"Schedule2 sheet, PA field empty r{row} c8");
                    WorkID = null;
                    return;
                }
                else
                {
                    InterviewSession = sValue2;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not read line:{row} in Schedule2. Error{ex.Message}");
            }

        }
    }
}

