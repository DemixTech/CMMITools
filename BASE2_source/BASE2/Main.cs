using BASE2.Data;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace BASE2
{
    public partial class Main : Form
    {
        #region globals

        public const string cPath_start = @"C:\Users\PietervanZyl\Demix (Pty) Ltd\Demix Global - PieterVZ\4_Appraisals\2020-12-11 (A5) R370 D5360 C51813 Goshine Tech";
        public const int cProjectHeadingStartRow = 2; // tab:Projects start row
        public const int cSupportHeadingStartRow = 2; // tab:Support start row
        public const int cStaffHeadingStartRow = 2; // tab:Staff start row
        public const int cSchedule2HeadingStartRow = 1; // tab:Schedule2 heading row

        public const int cOEDatabaseHeadingStartRow = 8; // All PAs
        public const int cOEnonEmptyColumn = 1; // 
        public const int cOEDatabaseMaxRows = 1000; // End of OE dbs

        private const int cXXSearchNumberOfWksRowsCol = 2;
        private const int cGOVandIIPASearchNumberOfWksRowsCol = 3; // Was 3

        private const int cMostPAStartRow = 2;
        private const int cMostPAEndRow = 20000;

        private const int cPAtestColumn = 1; // Column where the PA are listed for example GOV 2.3

        // private const int cMostPARowsTestCol = 1;
        private const int cMostPAtestOoS = 3;
        private const int cIIandGOVOoSTestCol = 4;


        private const int cXXWeaknessCol = 9; // col I
        private const int cIIandGOVWeaknessCol = 10;

        private const int cXXStrengthCol = 11; // col K
        private const int cIIandGOVStrenghtCol = 12;

        private const int cXXQuestionCol = 14; // Col N
        private const int cIIandGOVQuestionCol = 15; // Col O

        private const int cXXImprovementCol = 15; // col O
        private const int cIIandGOVImprovementCol = 16; // Col P

        // private const int cGOVrowCountTestCol = 1;
        //   private const int cIIrowCountTestCol = 2;


        private const int cIIMaxRows = 3607;
        private const int cGOVMaxRows = 4803;



        public string sourceFileName = "";

        // *** Private variables
        private List<WorkUnit> WorkUnitList = new List<WorkUnit>(); // Contain all the work unit detail

        // In a staff list, names can be duplicated accross multiople projects, but WorkID-Name will be unqiue
        private List<Staff> StaffList = new List<Staff>(); // Contian all the participant
        private List<OUProcess> OUProcessesList = new List<OUProcess>(); // Contain the processess and all their projects. Should only be initialised after WorkUnitList is established
        private List<Schedule2> Schedule2List = new List<Schedule2>(); // Contains the schedule 2 information
        private List<PracticeArea> CMMIModel = new List<PracticeArea>();


        //private Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
        private Workbook aWorkbook;
        private Workbook sourceWorkbook;
        private Workbook mainWorkbook;

        private Workbook questionWorkbook; // The workbook that contains the questions and the model

        public PersistentData persistentData = new PersistentData();
        #endregion

        public Main()
        {
            InitializeComponent();
        }

        //private void label3_Click(object sender, EventArgs e)
        //{

        //}

        private void btnSelectBASEplan_Click(object sender, EventArgs e)
        {
            #region btnSelectPlanTab

            // Clear background color
            //lbStatCASPlanLoaded.BackColor = Control.DefaultBackColor;

            // Check if the excel process is running

            OpenFileDialog sourceFile2 = new OpenFileDialog();
            sourceFile2.InitialDirectory = persistentData.LastAppraisalDirectory; //cPath_start;
            sourceFile2.RestoreDirectory = true;
            sourceFile2.Title = "Select source file";
            sourceFile2.DefaultExt = "*.xlsx";
            if (sourceFile2.ShowDialog() == DialogResult.OK)
            {
                // Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor;

                //  LblSourceFilePlan2.Text = sourceFile2.FileName;
                // *** Save and show selected plan file and directory
                persistentData.LastAppraisalDirectory = Path.GetDirectoryName(sourceFile2.FileName); // Path.GetDirectoryName(sourceFile2.FileName);
                persistentData.CASPlanName = Path.GetFileName(sourceFile2.FileName);

                // *** Store peristant data
                persistentData.SavePersistentData(persistentData);
                lblWorkingDirectoryText.Text = Path.GetDirectoryName(sourceFile2.FileName);
                lblWorkingFileText.Text = Path.GetFileName(sourceFile2.FileName);

                //lblWorkingDirectoryText.Text = persistentData.LastAppraisalDirectory;
                //lblPlanName.Text = persistentData.CASPlanName;

                //excelApp.Visible = true;

                //aWorkbook = excelApp.Workbooks.Open(LblSourceFilePlan2.Text.ToString());

                if ((aWorkbook = Helper.CheckIfOpenAndOpen(sourceFile2.FileName.ToString()))==null) // LblSourceFilePlan2.Text.ToString())) == null)
                {
                    //MessageBox.Show($"File {Path.GetFileName(LblSourceFilePlan2.Text.ToString())}" +
                    //    $"\n\rDirectory {Path.GetDirectoryName(LblSourceFilePlan2.Text.ToString())}" +
                    //    "\n\rdoes not exists");


                    // Set cursor as default arrow
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show("File not found, has it been moved or deleted?");
                    return;
                }

                // Step 0: Clear the list to start afresh
                WorkUnitList.Clear();
                StaffList.Clear();

                // Step 1: Open the spreadhseet and process it
                Worksheet projectWks = (Worksheet)aWorkbook.Sheets["Project&Support"];
                int row = cProjectHeadingStartRow + 1;
                Range aRng = (Range)projectWks.Cells[row,1];// .Rows[row],  1, "A"]; // .Range[]

                string sValue2 = aRng.Value.ToString(); // projectWks.Cells[row, 1].ToString();// .Value2;
                while (!string.IsNullOrEmpty(sValue2))
                {
                    // Process the list
                    WorkUnit aNewWorkUnitItem;
                    if (sValue2[0] == 'p')
                    {
                        aNewWorkUnitItem = new WorkUnit()
                        {
                            WorkType = EWorkType.project,
                        };
                        aNewWorkUnitItem.AddWorkType(EWorkType.project, projectWks, row, cProjectHeadingStartRow);

                    }
                    else
                    {
                        aNewWorkUnitItem = new WorkUnit()
                        {
                            WorkType = EWorkType.support,
                        };
                        aNewWorkUnitItem.AddWorkType(EWorkType.support, projectWks, row, cProjectHeadingStartRow);
                    }

                    if (aNewWorkUnitItem.WorkType != EWorkType.nothing) WorkUnitList.Add(aNewWorkUnitItem);

                    row++;
                    Range aRange2 = (Range)projectWks.Cells[row, 1];
                    sValue2 = aRange2?.Value?.ToString() ?? "";// projectWks.Cells[row, 1].ToString(); //.Value2;
                }
                // Step 2: Create the process list
                OUProcessesList.Clear();

                // Start at col 29 (AC) and search to the right until you find END
                int columnX = 29;
                int headerRow = 2; // Row where the processes are defined (below this row is the marking for the projects)
                string cellProcess = projectWks.Cells[headerRow, columnX].ToString();// .Value;
                int lastRowToProcess = Helper.FindEndOfWorksheet(projectWks, 1, 3, 50);
                while (cellProcess != "END")
                {
                    // Load the process name
                    OUProcess aProcess = new OUProcess();
                    aProcess.Name = cellProcess;

                    // Find asssociated projects
                    for (int rowX = 3; rowX <= lastRowToProcess; rowX++)
                    {
                        string cellMarkedX = projectWks.Cells[rowX, columnX]?.ToString(); //.Value;
                        if (cellMarkedX?.ToLower() == "x")
                        { // Marked x, proxcess it
                            string workIdStr = projectWks.Cells[rowX, 1]?.ToString(); //.Value;
                            // The workId must be valid, cannot be null or empty
                            if (string.IsNullOrEmpty(workIdStr))
                            {
                                MessageBox.Show($"WorkID at {rowX} cannot be null or empty");
                            }
                            else
                            {
                                // Use the workIdStr to find the WorkUnit and attach it to the process
                                WorkUnit aWorkunit = WorkUnitList.Find(x => x.ID == workIdStr);
                                if (aWorkunit == null)
                                {
                                    MessageBox.Show($"No work unit found in list for {workIdStr}! Please review Projects table.");
                                }
                                else
                                { // Add the work unit found
                                    aProcess.WorkUnits.Add(aWorkunit);
                                }
                            }

                        }
                    }
                    // Add the process and search for the next one in the next column
                    OUProcessesList.Add(aProcess);

                    // Test for endless loop
                    if (columnX++ > 100)
                    {
                        MessageBox.Show("END not found. See if end is listed in Row 2 of Projects tab!");
                        break;
                    }
                    else
                    {
                        cellProcess = projectWks.Cells[headerRow, columnX].ToString();//.Value;
                    }
                } // Process until end is found

                // Step 2: Open the support spreadhseet and process it
                //Worksheet supportWks = aWorkbook.Sheets["Support"];
                //row = cSupportHeadingStartRow + 1;
                //string sValue4 = supportWks.Cells[row, 1].Value2;
                //while (!string.IsNullOrEmpty(sValue4))
                //{
                //    // Process the list
                //    WorkUnit aNewWorkUnitItem = new WorkUnit(EWorkType.support, supportWks, row, cSupportHeadingStartRow);
                //    if (aNewWorkUnitItem.WorkType != EWorkType.nothing) WorkUnitList.Add(aNewWorkUnitItem);

                //    row++;
                //    sValue4 = supportWks.Cells[row, 1].Value2;
                //}

                // Step 3: Open the participant spreadhseet and process it
                Worksheet participantWks = (Worksheet)aWorkbook.Sheets["Staff"];
                row = cStaffHeadingStartRow + 1;
                string sValue5 = participantWks.Cells[row, 1]?.ToString(); //.Value2;
                while (!string.IsNullOrEmpty(sValue5))
                {
                    // Process the list
                    Staff aNewParticipant = new Staff();
                    aNewParticipant.StaffAdd(participantWks, row, cStaffHeadingStartRow);
                    if (aNewParticipant.WorkID != null) StaffList.Add(aNewParticipant);

                    row++;
                    sValue5 = participantWks.Cells[row, 1].ToString(); // .Value2;
                }
                // *** Load OU information

                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Workbook loaded. Projects and support functions loaded. Processess loaded. Staff loaded.");

                // Step 4: Load Scheduel 2
                /// ************* loadSchedule2();
            }

            // Set background color - loaded
            //lbStatCASPlanLoaded.BackColor = Color.LightGreen;

            #endregion

        }
    }
}
