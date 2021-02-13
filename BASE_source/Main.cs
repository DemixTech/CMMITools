/*
    Copyright (c) 2020-2021, Demix (Pty) Ltd, Create | Evolve | Perfect, http://www.demix.org
    License agreement https://github.com/DemixTech/CMMITools/blob/main/README.md
*/


using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using BASE.Data;
using System.Reflection;
using System.IO.Compression;

//using Microsoft.Office.Interop.Excel;

namespace BASE
{
    public partial class Main : Form
    {

        #region globals
        const string CTargetCASFileXML = @"BASE\TargetCASFileXML.xml";
        const string CTargetOEdbFileXML = @"BASE\TargetOEdbFileXML.xml";
        const string CTargetOEdbImportFileXML = @"BASE\TargetOEdbImportFileXML.xml";
        const string CQuestionFileXML = @"BASE\TargetQuestionModelFileXML.xml";
        const string CDataReferenceFileXML = @"BASE\TargetDataReferenceFileXML.xml";
        const string CPresentationFileXML = @"BASE\TargetPresentationFileXML.xml";

        private const string cPath_start = @"C:\Users\PietervanZyl\Demix (Pty) Ltd\Demix Global - PieterVZ\4_Appraisals\2020-12-11 (A5) R370 D5360 C51813 Goshine Tech";
        private const int cProjectHeadingStartRow = 2; // tab:Projects start row
        private const int cSupportHeadingStartRow = 2; // tab:Support start row
        private const int cStaffHeadingStartRow = 2; // tab:Staff start row
        private const int cSchedule2HeadingStartRow = 1; // tab:Schedule2 heading row

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

        //    private Workbook questionWorkbook; // The workbook that contains the questions and the model

        public PersistentData persistentData = new PersistentData();

        // *** BASE file objects
        private TargetCASFileObject CASFileObject;
        private TargetOEFileObject CASOEdbObject;
        private TargetOEFileObject CASOEdbImportObject;
        private TargetQuestionsFileObject BASEQuestionObject;

        private TargetDataReferenceFileObject BASEDataReferenceObject;
        private TargetPresentationFileObject BASEPresentationObject;



        #endregion

        public string VersionLabel
        {
            get
            {
                if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
                {
                    Version ver = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
                    return string.Format("Product Name: {4}, Version: {0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision, Assembly.GetEntryAssembly().GetName().Name);
                }
                else
                {
                    var ver = Assembly.GetExecutingAssembly().GetName().Version;
                    return string.Format("Product Name: {4}, Version: {0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision, Assembly.GetEntryAssembly().GetName().Name);
                }
            }
        }

        public Main()
        {
            InitializeComponent();
            // lblWorkingDirectory.Text = Directory.GetCurrentDirectory();
            //lblWorkingDirectory.Text = Path.GetTempPath(); //Environment.SpecialFolder.Personal.ToString();

            // *** Startup program objects
            CASFileObject = new TargetCASFileObject();
            CASFileObject.InitialiseObject(Path.Combine(Path.GetTempPath(), CTargetCASFileXML), lblCASPathXML, lblCASFileXML, lblCASPlanPathText, lblCASPlanFileText);
            CASFileObject.LoadPersistantXMLdata(); // 

            // *** Startup oeDb objects
            CASOEdbObject = new TargetOEFileObject();
            CASOEdbObject.InitialiseObject(Path.Combine(Path.GetTempPath(), CTargetOEdbFileXML), lblOEPathXML2, lblOEFileXML2, lblOEPath2, lblOEFile2);
            CASOEdbObject.LoadPersistantXMLdata(); // 

            // *** Startup oeDb import objects
            CASOEdbImportObject = new TargetOEFileObject();
            CASOEdbImportObject.InitialiseObject(Path.Combine(Path.GetTempPath(), CTargetOEdbImportFileXML), lblPathOEdbImportXML2, lblFileOEdbImportXML2,
                lblPathOEdbImport2, lblFileOEdbImport2); ;
            CASOEdbImportObject.LoadPersistantXMLdata(); //

            // *** Startup Question objects
            BASEQuestionObject = new TargetQuestionsFileObject();
            BASEQuestionObject.InitialiseObject(Path.Combine(Path.GetTempPath(), CQuestionFileXML), lblQMPathXML2, lblQMfileXML2, lblQuestionPath2, lblQuestionFile2);
            BASEQuestionObject.LoadPersistantXMLdata(); // 

            // *** Startup Data reference object
            BASEDataReferenceObject = new TargetDataReferenceFileObject();
            BASEDataReferenceObject.InitialiseObject(CDataReferenceFileXML, lblDataReferenceXMLPath2, lblDataReferenceXMLFile2,
                lblXlsxPath2, lblXlsxFile2);
            BASEDataReferenceObject.LoadPersistantXMLdata();

            // *** Startup presentation object
            BASEPresentationObject = new TargetPresentationFileObject();
            BASEPresentationObject.InitialiseObject(CPresentationFileXML, lblPresentationXMLPath2, lblPresentationXMLFile2,
                lblPptxPath2, lblPptxFile2);
            BASEPresentationObject.LoadPersistantXMLdata();


            // *** Old code
            persistentData.LoadPersistentData();
            lblWorkingDirectory.Text = persistentData.LastAppraisalDirectory;
            lblPlanName.Text = persistentData.CASPlanName;

            lblDefaults.Text = PersistentData.PersistantPathFile_Generic;  // persistentData.PersistantPathFile_Generic;
            lblOEdbMain.Text = persistentData.AppToolMainPathFile;
            lblOEdbSource.Text = persistentData.AppToolSourcePathFile;
            lblOEdbPathFile.Text = persistentData.OEdatabasePathFile;
            //lblDemixTool.Text = persistentData.DemixToolPathFile;
            //lblDemixTool2Import.Text = persistentData.DemixTool_ToImport_PathFile;

            //lblQuestions.Text = persistentData.QuestionPathFile;

            txtFrom.Text = persistentData.FromText;
            txtTo.Text = persistentData.ToText;

            PersistentData.LoadPersistentData_Questions(ref CMMIModel);
            PersistentData.LoadPersistentData_WorkUnitList(ref WorkUnitList);
            PersistentData.LoadPersistentData_ProcessList(ref OUProcessesList);
            PersistentData.LoadPersistentData_StaffList(ref StaffList);
            PersistentData.LoadPersistentData_Schedule2List(ref Schedule2List);

        }

        private void btnSelectPlan_Click(object sender, EventArgs e)
        {

        }

        private void buttonGenerateSchedule_Click(object sender, EventArgs e)
        {

        }

        private void Main_Load(object sender, EventArgs e)
        {
            var assemblyVersion = System.Windows.Forms.Application.ProductVersion; // comment out //[assembly: AssemblyFileVersion("1.0.0.0")] in AssemblyInfo.cs
                                                                                   //var fileVersion = System.Reflection.AssemblyFileVersionAttribute.
                                                                                   //string assemblyVersion2;// = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
                                                                                   //try
                                                                                   //{
                                                                                   //    assemblyVersion2 = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
                                                                                   //}
                                                                                   //catch (Exception ex)
                                                                                   //{
                                                                                   //    assemblyVersion2 = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                                                                                   //}
                                                                                   //  MessageBox.Show($"Assembly1 {assemblyVersion}");//\nAssembly2 {assemblyVersion2.ToString()}");

            string headerText = this.Text;
            headerText = headerText + $" version {assemblyVersion} - Copyright(c) 2020, 2021 Demix (Pty) Ltd, All rights reserved";
            this.Text = headerText;
        }


        private void btnLoadXML_Click(object sender, EventArgs e)
        {
            persistentData.LoadPersistentData();
            lblWorkingDirectory.Text = persistentData.LastAppraisalDirectory;
        }

        private void btnSaveXML_Click(object sender, EventArgs e)
        {
            persistentData.SavePersistentData(persistentData);
        }

        private void Main_FormClosed(object sender, FormClosedEventArgs e)
        {

            persistentData.SavePersistentData(persistentData);
            PersistentData.SavePersistentData_Questions(CMMIModel);
            PersistentData.SavePersistentData_WorkUnits(WorkUnitList);
            PersistentData.SavePersistentData_ProcessLists(OUProcessesList);
            PersistentData.SavePersistentData_StaffList(StaffList);
            PersistentData.SavePersistentData_Schedule2List(Schedule2List);
        }

        private void btnOEdbMain_Click(object sender, EventArgs e)
        {

        }

        private void btnOEdbSource_Click(object sender, EventArgs e)
        {

        }

        private void btnMergeSourceToMain_Click(object sender, EventArgs e)
        {


        }

        private void btnSetupMain_Click(object sender, EventArgs e)
        {

        }

        private void btnLoadSchedule2_Click(object sender, EventArgs e)
        {

        }

        // *** Helper functions
        private void loadSchedule2()
        {
            string filePath = Path.Combine(lblWorkingDirectory.Text, lblPlanName.Text);
            //excelApp.Visible = true;
            // aWorkbook = excelApp.Workbooks.Open(filePath);


            if ((aWorkbook = Helper.CheckIfOpenAndOpenXlsx(filePath)) == null)
            {
                //MessageBox.Show($"File {Path.GetFileName(filePath)}" +
                //    $"\n\rDirectory {Path.GetDirectoryName(filePath)}" +
                //    "\n\rdoes not exists");
                MessageBox.Show("File not found, has it been moved or deleted?");
                return;
            }


            // Step 1: Open the spreadhseet and process it
            Schedule2List.Clear();
            Worksheet projectWks = aWorkbook.Sheets["Schedule2"];
            int NumberOfRows = Helper.FindEndOfWorksheet(projectWks, 1, cSchedule2HeadingStartRow, 500);
            for (int row = cSchedule2HeadingStartRow + 1; row <= NumberOfRows; row++)
            {
                // Process the Schedule 2 list
                Schedule2 aNewSchedule2Record = new Schedule2();
                aNewSchedule2Record.Schedule2Add(projectWks, row);

                if (aNewSchedule2Record.WorkID != null) Schedule2List.Add(aNewSchedule2Record);
            }
            MessageBox.Show($"Schedule2 loaded. Records={NumberOfRows - cSchedule2HeadingStartRow}");
        }

        private string selectInterviewAndName(string PAname)
        {
            //wksMain.Name);// "[x] Interview - Name"
            string InterviewAndName = "";
            List<Schedule2> subSetSchedule2 = Schedule2List.Where(x => x.PA == PAname).ToList();
            foreach (Schedule2 schedule2Record in subSetSchedule2)
            {
                if (InterviewAndName == "")
                { // First pass
                    InterviewAndName = schedule2Record.InterviewSession + " (" + schedule2Record.ParticipantName + ")";
                }
                else
                {
                    InterviewAndName = InterviewAndName + "\r\n" + schedule2Record.InterviewSession + " (" + schedule2Record.ParticipantName + ")";

                }
            }
            return InterviewAndName;
        }

        private void btnOEdb_Click(object sender, EventArgs e)
        {

        }

        private void btnTestLinkEngl_Click(object sender, EventArgs e)
        {




        }

        private void btnInsertInterviews_Click(object sender, EventArgs e)
        {


        }

        private void btnHideOoSRows_Click(object sender, EventArgs e)
        {

        }

        const int cPraciceNumberColumnIIandGOV = 1;
        const int cProcessNameColumnIIandGOV = 2;
        const int cProjectNameColumnIIandGOV = 3;


        // **** Helper function
        private void HelperIIGOVcharacterization(Worksheet wksMain, Worksheet wksIIGOVchar, int searchColumn, int startRow, int EndRow, int testColumnPA, int testColumnOoS)//, ref int SPOSFratingRow)
        {

            // *** Find the number of rows
            int NumberOfRows = Helper.FindEndOfWorksheet(wksMain, searchColumn, startRow, EndRow);
            Range mainRange = wksMain.Range["A" + startRow, "Z" + NumberOfRows];
            object[,] mainValue = mainRange.Value;

            // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
            //Range mainRange = wksMain.Range["A" + startRow, "E" + NumberOfRows];
            //Range scopeRange = wksMain.Range["C" + 2, "C" + NumberOfRows];
            //string practiceNumber = "Unassigned";
            string processName = "Unassigned";
            string projectName = "Unassinged";
            string practiceNumberStr = "Unassigned";

            for (int rowS = 1; rowS <= NumberOfRows - startRow + 1; rowS++)
            {
                //string cellA = mainRange[rowS, cPraciceNumberColumnIIandGOV]?.Value?.ToString();
                //if (!string.IsNullOrEmpty(cellA))
                //{
                //    practiceNumber = cellA;
                //}
                string cellA = mainRange[rowS, cPraciceNumberColumnIIandGOV]?.Value?.ToString();
                if (!string.IsNullOrEmpty(cellA)) practiceNumberStr = cellA;

                string cellB = mainRange[rowS, cProcessNameColumnIIandGOV]?.Value?.ToString();
                if (!string.IsNullOrEmpty(cellB)) processName = cellB;

                string cellC = mainRange[rowS, cProjectNameColumnIIandGOV]?.Value?.ToString();
                if (!string.IsNullOrEmpty(cellC)) projectName = cellC;

                //bool hiddenRow = true; // Default hide it
                string cell1 = mainRange[rowS, testColumnPA]?.Value?.ToString();
                string cell2 = mainRange[rowS, testColumnOoS]?.Value?.ToString();
                if (cell2 != "OoS")
                {

                    int theRow = findArowInSPOSFrating(wksIIGOVchar, processName, projectName);
                    // Write this entry
                    wksIIGOVchar.Cells[theRow, 1].Value = processName;
                    wksIIGOVchar.Cells[theRow, 2].Value = projectName;
                    wksIIGOVchar.Cells[theRow, ColumnOfPractice(practiceNumberStr)].Value = cell2;
                    // wksIIGOVchar.Cells[SPOSFratingRow, 3].Value = practiceNumber;
                    //  wksIIGOVchar.Cells[SPOSFratingRow, 4].Value = cell2;
                    //SPOSFratingRow++;
                    //hiddenRow = false; // Show if null and NOT OoS

                }
                //if (cell1 != null) hiddenRow = false; // Show if not null
                //mainRange.Rows[rowS].EntirRow.Hidden = hiddenRow; // .HideRow(rowS);
                //if (hiddenRow) wksMain.Rows[rowS].EntireRow.Hidden = true; // .Height = 0; //.Hide(); // .Unhide();
                //else wksMain.Rows[rowS].EntireRow.Hidden = false; // .Unhide();
                //Range rowRange = wksMain.Rows[rowS + startRow - 1]; // mainRange[rowS,1].EntireRow.Hidden = hiddenRow;
                //rowRange.Hidden = hiddenRow;
            }
        }

        const int cSPOSFratingStrtRow = 4;
        const int cSPOSFratingMaxRows = 100;

        private int findArowInSPOSFrating(Worksheet wksIIGOVchar, string processName, string projectName)
        {
            // int theFinalRow = cSPOSFratingStrtRow;

            processName = processName.ToLower().Trim();
            projectName = projectName.ToLower().Trim();
            // *** find the end of the sheets

            // Start from heading row. If now data, then number of rows will be less than cSPOSFratingSrtRow
            int NumberOfRows = Helper.FindEndOfWorksheet(wksIIGOVchar, 1, cSPOSFratingStrtRow - 1, cSPOSFratingMaxRows);
            if (NumberOfRows < cSPOSFratingStrtRow) return cSPOSFratingStrtRow; // Start from here

            // else number of rows is th elast item

            Range mainRange = wksIIGOVchar.Range["A" + cSPOSFratingStrtRow, "B" + NumberOfRows];
            object[,] mainValue = mainRange.Value;

            for (int i = 1; i <= NumberOfRows - cSPOSFratingStrtRow + 1; i++)
            {
                if (mainValue[i, 1]?.ToString().ToLower().Trim() == processName && mainValue[i, 2]?.ToString().ToLower().Trim() == projectName)
                {
                    return i + cSPOSFratingStrtRow - 1;
                }
            }
            return NumberOfRows + 1;

        }

        private void HelperHideRows2(Worksheet wksMain, int searchColumn, int startRow, int EndRow, int testColumnPA, int testColumnOoS)
        {

            // *** Find the number of rows
            int NumberOfRows = Helper.FindEndOfWorksheet(wksMain, searchColumn, startRow, EndRow);
            Range mainRange = wksMain.Range["A" + startRow, "Z" + NumberOfRows];
            object[,] mainValue = mainRange.Value;

            // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
            //Range mainRange = wksMain.Range["A" + startRow, "E" + NumberOfRows];
            //Range scopeRange = wksMain.Range["C" + 2, "C" + NumberOfRows];

            for (int rowS = 1; rowS <= NumberOfRows - startRow + 1; rowS++)
            {
                bool hiddenRow = true; // Default hide it
                string cell1 = mainRange[rowS, testColumnPA]?.Value?.ToString();
                string cell2 = mainRange[rowS, testColumnOoS]?.Value?.ToString();
                if (cell1 == null && cell2 != "OoS") hiddenRow = false; // Show if null and NOT OoS
                if (cell1 != null) hiddenRow = false; // Show if not null
                //mainRange.Rows[rowS].EntirRow.Hidden = hiddenRow; // .HideRow(rowS);
                //if (hiddenRow) wksMain.Rows[rowS].EntireRow.Hidden = true; // .Height = 0; //.Hide(); // .Unhide();
                //else wksMain.Rows[rowS].EntireRow.Hidden = false; // .Unhide();
                Range rowRange = wksMain.Rows[rowS + startRow - 1]; // mainRange[rowS,1].EntireRow.Hidden = hiddenRow;
                rowRange.Hidden = hiddenRow;
            }
        }

        private void HelperHideRows1(Worksheet wksMain, int searchColumn, int startRow, int EndRow, int testColumnPA, int testColumnOoS)
        {

            // *** Find the number of rows
            int NumberOfRows = Helper.FindEndOfWorksheet(wksMain, searchColumn, startRow, EndRow);

            // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
            //Range mainRange = wksMain.Range["A" + startRow, "E" + NumberOfRows];
            //Range scopeRange = wksMain.Range["C" + 2, "C" + NumberOfRows];
            for (int rowS = startRow; rowS <= NumberOfRows; rowS++)
            {
                bool hiddenRow = true; // Default hide it
                string cell1 = wksMain.Cells[rowS, testColumnPA]?.Value?.ToString();
                string cell2 = wksMain.Cells[rowS, testColumnOoS]?.Value?.ToString();
                if (cell1 == null && cell2 != "OoS") hiddenRow = false; // Show if null and NOT OoS
                if (cell1 != null) hiddenRow = false; // Show if not null
                //mainRange.Rows[rowS].EntirRow.Hidden = hiddenRow; // .HideRow(rowS);
                //if (hiddenRow) wksMain.Rows[rowS].EntireRow.Hidden = true; // .Height = 0; //.Hide(); // .Unhide();
                //else wksMain.Rows[rowS].EntireRow.Hidden = false; // .Unhide();
                Range rowRange = wksMain.Rows[rowS]; // mainRange[rowS,1].EntireRow.Hidden = hiddenRow;
                rowRange.Hidden = hiddenRow;
            }
        }

        private void btnLoadSchedule2_tab_Click(object sender, EventArgs e)
        {

        }

        private void btnLoadSchedule2tab_Click(object sender, EventArgs e)
        {
            // Remove from release 3.8.0.44
            //// loadSchedule2();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void btnGenerateScheduleTab_Click(object sender, EventArgs e)
        {
            ////// Remove from release 3.8.0.44
            ////// Step 4: Show schedule
            ////Worksheet schedule = aWorkbook.Sheets["Schedule"];
            ////schedule.Cells.Clear();
            ////schedule.Cells[1, 1].Value = "WorkID";
            ////schedule.Cells[1, 2].Value = "Work name";
            ////schedule.Cells[1, 3].Value = "PA";
            ////schedule.Cells[1, 4].Value = "Participant Name";
            ////schedule.Cells[1, 5].Value = "Role";
            ////schedule.Cells[1, 6].Value = "WordID2";
            ////schedule.Cells[1, 7].Value = "Included";
            ////// https://www.e-iceblue.com/Tutorials/Spire.XLS/Spire.XLS-Program-Guide/Set-Excel-Background-Color-with-C-VB.NET.html
            ////// schedule.Range["A1:A6"].Style.Color = Color.BlueViolet;

            ////// *** For each project selected PA, find the participants that acted in that role
            ////int outRow = 2;
            ////List<Schedule1Entry> includedList = new List<Schedule1Entry>();
            ////List<Schedule1Entry> excludedList = new List<Schedule1Entry>();

            ////foreach (var workUnit in WorkUnitList)
            ////{
            ////    var listOfSampledPAs = workUnit.PAlist.Where(x => x.SampleType == ESampleType.added || x.SampleType == ESampleType.sampled);
            ////    foreach (var aSampledPA in listOfSampledPAs)
            ////    {
            ////        // This is all the sampled PAs
            ////        var workUnitStaffList = StaffList.Where(x => x.WorkID == workUnit.ID);
            ////        foreach (var workUnitParticipant in workUnitStaffList)
            ////        {
            ////            // Find the practice areas that match
            ////            var participantForSampledWorkUnitPA = workUnitParticipant.PAlist.Where(x => x.PAcode == aSampledPA.PAcode);
            ////            bool found;
            ////            if (participantForSampledWorkUnitPA.Count() == 0)
            ////            { // nothing found here
            ////                found = false; // nothing found
            ////            }
            ////            else
            ////            {
            ////                // proces this list
            ////                found = true;
            ////                schedule.Cells[outRow, 1].Value = workUnit.ID.ToString();
            ////                schedule.Cells[outRow, 2].Value = workUnit.Name.ToString();
            ////                schedule.Cells[outRow, 3].Value = aSampledPA.PAcode.ToString();
            ////                schedule.Cells[outRow, 4].Value = workUnitParticipant.Name;
            ////                schedule.Cells[outRow, 5].Value = workUnitParticipant.Role;
            ////                schedule.Cells[outRow, 6].Value = workUnitParticipant.WorkID;

            ////                // check if the workUnit.ID and PAcode is not already in the list, if it is, then it is a unecessary duplicate. If duplicate, skip include
            ////                Schedule1Entry aSchedule1Entry = new Schedule1Entry()
            ////                {
            ////                    ID = workUnit.ID.ToString(),
            ////                    Name = workUnit.Name.ToString(),
            ////                    PAcode = aSampledPA.PAcode.ToString(),
            ////                    ParticipantName = workUnitParticipant.Name,
            ////                    ParticipantRole = workUnitParticipant.Role,
            ////                    WorkIDcheck = workUnitParticipant.WorkID,
            ////                    //  include = false, set below to make reading more clear
            ////                };

            ////                int inListIndex = includedList.FindIndex(x => x.ID == aSchedule1Entry.ID && x.PAcode == aSchedule1Entry.PAcode); ;
            ////                if (inListIndex < 0)
            ////                { // not in list, insert in includedList
            ////                    aSchedule1Entry.include = true;
            ////                    includedList.Insert(~inListIndex, aSchedule1Entry);
            ////                    schedule.Cells[outRow, 7].Value = "x";
            ////                }
            ////                else
            ////                { // already included, put in excludedList
            ////                    aSchedule1Entry.include = false;
            ////                    excludedList.Add(aSchedule1Entry);
            ////                    schedule.Cells[outRow, 7].Value = "";
            ////                }
            ////                outRow++;


            ////            }
            ////        }
            ////    }
            ////}

            ////// *** Find distinct participants
            ////Worksheet responsibilities = aWorkbook.Sheets["Responsibilities"];
            ////responsibilities.Cells.Clear();
            ////int respRow = 2;
            ////var distinctParticipants = StaffList.Select(x => x.Name)
            ////    .Distinct()
            ////    .OrderBy(q => q)
            ////    .ToList();

            ////List<string> projectNameList = new List<string>();
            ////List<string> projectWorkIDList = new List<string>();
            ////List<EPAcode> practiceAreaList = new List<EPAcode>();
            ////foreach (var distincParticipant in distinctParticipants)
            ////{
            ////    // first clear the names and practice list
            ////    projectNameList.Clear();
            ////    projectWorkIDList.Clear();
            ////    practiceAreaList.Clear();

            ////    // *** List all the Projects
            ////    var participantSubset = StaffList.Where(x => x.Name == distincParticipant);
            ////    foreach (var aParticipant in participantSubset)
            ////    {
            ////        // Add a project if it does not exists
            ////        var x = projectNameList.BinarySearch(aParticipant.WorkName);
            ////        if (x < 0)
            ////        { // Not in list, add it
            ////            projectNameList.Insert(~x, aParticipant.WorkName);
            ////            projectWorkIDList.Insert(~x, aParticipant.WorkID);
            ////        }
            ////        else
            ////        {
            ////            // in list, ignore it
            ////        }
            ////        // Add the PAs if it does not exist
            ////        foreach (var aPa in aParticipant.PAlist)
            ////        {
            ////            var pai = practiceAreaList.BinarySearch(aPa.PAcode);
            ////            if (pai < 0)
            ////            {
            ////                // Not in list, add it
            ////                practiceAreaList.Insert(~pai, aPa.PAcode);
            ////            }
            ////            else
            ////            {
            ////                // In list, ignore it
            ////            }
            ////        }
            ////    }

            ////    // for this participant, output the practicenames and practice areas to the spreadhseet
            ////    responsibilities.Cells[respRow++, 1].Value = distincParticipant;
            ////    responsibilities.Cells[respRow++, 2].Value = "Project/Work";
            ////    for (int i = 0; i < projectNameList.Count(); i++)
            ////    //foreach (var aprojName in projectNameList)
            ////    {
            ////        //responsibilities.Cells[respRow++, 3].Value = aprojName.ToString();
            ////        // projectWorkIDList
            ////        responsibilities.Cells[respRow++, 3].Value =
            ////            projectWorkIDList[i] + " " + projectNameList[i];
            ////        //aprojName.ToString();
            ////    }
            ////    responsibilities.Cells[respRow++, 2].Value = "Practice Area";
            ////    foreach (var aPA in practiceAreaList)
            ////    {
            ////        responsibilities.Cells[respRow++, 3].Value = aPA.ToString();
            ////    }

            ////}

            ////MessageBox.Show("Draft Schedule completed");

        }

        private void btnSelectPlanTab_Click(object sender, EventArgs e)
        {
            // Delete from release 3.8.0.44

            ////#region btnSelectPlanTab

            ////// Clear background color
            //////lblstat lbStatCASPlanLoaded.BackColor = Control.DefaultBackColor;

            ////// Check if the excel process is running

            ////OpenFileDialog sourceFile2 = new OpenFileDialog();
            ////sourceFile2.InitialDirectory = persistentData.LastAppraisalDirectory; //cPath_start;
            ////sourceFile2.RestoreDirectory = true;
            ////sourceFile2.Title = "Select source file";
            ////sourceFile2.DefaultExt = "*.xlsx";
            ////if (sourceFile2.ShowDialog() == DialogResult.OK)
            ////{
            ////    // Set cursor as hourglass
            ////    Cursor.Current = Cursors.WaitCursor;

            ////    LblSourceFilePlan2.Text = sourceFile2.FileName;
            ////    // *** save new file
            ////    persistentData.LastAppraisalDirectory = Path.GetDirectoryName(sourceFile2.FileName);
            ////    persistentData.CASPlanName = Path.GetFileName(sourceFile2.FileName);
            ////    persistentData.SavePersistentData(persistentData);
            ////    lblWorkingDirectory.Text = persistentData.LastAppraisalDirectory;
            ////    lblPlanName.Text = persistentData.CASPlanName;

            ////    //excelApp.Visible = true;

            ////    //aWorkbook = excelApp.Workbooks.Open(LblSourceFilePlan2.Text.ToString());

            ////    if ((aWorkbook = Helper.CheckIfOpenAndOpen(LblSourceFilePlan2.Text.ToString())) == null)
            ////    {
            ////        //MessageBox.Show($"File {Path.GetFileName(LblSourceFilePlan2.Text.ToString())}" +
            ////        //    $"\n\rDirectory {Path.GetDirectoryName(LblSourceFilePlan2.Text.ToString())}" +
            ////        //    "\n\rdoes not exists");


            ////        // Set cursor as default arrow
            ////        Cursor.Current = Cursors.Default;
            ////        MessageBox.Show("File not found, has it been moved or deleted?");
            ////        return;
            ////    }

            ////    // Step 0: Clear the list to start afresh
            ////    WorkUnitList.Clear();
            ////    StaffList.Clear();

            ////    // Step 1: Open the spreadhseet and process it
            ////    Worksheet projectWks = aWorkbook.Sheets["Project&Support"];
            ////    int row = cProjectHeadingStartRow + 1;
            ////    string sValue2 = projectWks.Cells[row, 1].Value2;
            ////    while (!string.IsNullOrEmpty(sValue2))
            ////    {
            ////        // Process the list
            ////        WorkUnit aNewWorkUnitItem;
            ////        char firstChar = sValue2.ToUpper()[0];
            ////        switch (firstChar)
            ////        {
            ////            case 'P':
            ////                aNewWorkUnitItem = new WorkUnit()
            ////                {
            ////                    WorkType = EWorkType.project,
            ////                };
            ////                aNewWorkUnitItem.AddWorkType(EWorkType.project, projectWks, row, cProjectHeadingStartRow);
            ////                break;
            ////            case 'S':
            ////                aNewWorkUnitItem = new WorkUnit()
            ////                {
            ////                    WorkType = EWorkType.support,
            ////                };
            ////                aNewWorkUnitItem.AddWorkType(EWorkType.support, projectWks, row, cProjectHeadingStartRow);
            ////                break;
            ////            default:
            ////                aNewWorkUnitItem = new WorkUnit()
            ////                {
            ////                    WorkType = EWorkType.nothing,
            ////                };
            ////                break;
            ////        }

            ////        WorkUnitList.Add(aNewWorkUnitItem);
            ////        row++;
            ////        sValue2 = projectWks.Cells[row, 1].Value2;
            ////    }
            ////    // Step 2: Create the process list
            ////    OUProcessesList.Clear();

            ////    // Start at col 29 (AC) and search to the right until you find END
            ////    int columnX = 29;
            ////    int headerRow = 2; // Row where the processes are defined (below this row is the marking for the projects)
            ////    string cellProcess = projectWks.Cells[headerRow, columnX].Value;
            ////    int lastRowToProcess = Helper.FindEndOfWorksheet(projectWks, 1, 3, 50);
            ////    while (cellProcess != "END")
            ////    {
            ////        // Load the process name
            ////        OUProcess aProcess = new OUProcess();
            ////        aProcess.Name = cellProcess;

            ////        // Find asssociated projects
            ////        for (int rowX = 3; rowX <= lastRowToProcess; rowX++)
            ////        {
            ////            string cellMarkedX = projectWks.Cells[rowX, columnX]?.Value;
            ////            if (cellMarkedX?.ToLower() == "x")
            ////            { // Marked x, proxcess it
            ////                string workIdStr = projectWks.Cells[rowX, 1]?.Value;
            ////                // The workId must be valid, cannot be null or empty
            ////                if (string.IsNullOrEmpty(workIdStr))
            ////                {
            ////                    MessageBox.Show($"WorkID at {rowX} cannot be null or empty");
            ////                }
            ////                else
            ////                {
            ////                    // Use the workIdStr to find the WorkUnit and attach it to the process
            ////                    WorkUnit aWorkunit = WorkUnitList.Find(x => x.ID.ToLower() == workIdStr.ToLower());
            ////                    if (aWorkunit == null)
            ////                    {
            ////                        MessageBox.Show($"No work unit found in list for {workIdStr}! Please review Projects table.");
            ////                    }
            ////                    else
            ////                    { // Add the work unit found
            ////                        aProcess.WorkUnits.Add(aWorkunit);
            ////                    }
            ////                }

            ////            }
            ////        }
            ////        // Add the process and search for the next one in the next column
            ////        OUProcessesList.Add(aProcess);

            ////        // Test for endless loop
            ////        if (columnX++ > 100)
            ////        {
            ////            MessageBox.Show("END not found. See if end is listed in Row 2 of Projects tab!");
            ////            break;
            ////        }
            ////        else
            ////        {
            ////            cellProcess = projectWks.Cells[headerRow, columnX].Value;
            ////        }
            ////    } // Process until end is found

            ////    // Step 2: Open the support spreadhseet and process it
            ////    //Worksheet supportWks = aWorkbook.Sheets["Support"];
            ////    //row = cSupportHeadingStartRow + 1;
            ////    //string sValue4 = supportWks.Cells[row, 1].Value2;
            ////    //while (!string.IsNullOrEmpty(sValue4))
            ////    //{
            ////    //    // Process the list
            ////    //    WorkUnit aNewWorkUnitItem = new WorkUnit(EWorkType.support, supportWks, row, cSupportHeadingStartRow);
            ////    //    if (aNewWorkUnitItem.WorkType != EWorkType.nothing) WorkUnitList.Add(aNewWorkUnitItem);

            ////    //    row++;
            ////    //    sValue4 = supportWks.Cells[row, 1].Value2;
            ////    //}

            ////    // Step 3: Open the participant spreadhseet and process it
            ////    Worksheet participantWks = aWorkbook.Sheets["Staff"];
            ////    row = cStaffHeadingStartRow + 1;
            ////    string sValue5 = participantWks.Cells[row, 1].Value2;
            ////    while (!string.IsNullOrEmpty(sValue5))
            ////    {
            ////        // Process the list
            ////        Staff aNewParticipant = new Staff();
            ////        aNewParticipant.StaffAdd(participantWks, row, cStaffHeadingStartRow);
            ////        if (aNewParticipant.WorkID != null) StaffList.Add(aNewParticipant);

            ////        row++;
            ////        sValue5 = participantWks.Cells[row, 1].Value2;
            ////    }
            ////    // *** Load OU information

            ////    // Set cursor as default arrow
            ////    Cursor.Current = Cursors.Default;
            ////    MessageBox.Show("Workbook loaded. Projects and support functions loaded. Processess loaded. Staff loaded.");

            ////    // Step 4: Load Scheduel 2
            ////    loadSchedule2();
            ////}

            ////// Set background color - loaded
            //////   lbStatCASPlanLoaded.BackColor = Color.LightGreen;

            ////#endregion

        }

        //private void button1_Click_1(object sender, EventArgs e)
        //{

        //}

        private void btnExcelRunning_Click(object sender, EventArgs e)
        {
            // Microsoft.Office.Interop.Excel.Application excelApp2 = new Microsoft.Office.Interop.Excel.Application();
            // https://social.msdn.microsoft.com/Forums/vstudio/en-US/c2a48936-b58f-4487-84ec-dfae842e2fd1/how-to-check-to-see-if-excel-application-is-already-open?forum=csharpgeneral
            Process[] processlist = Process.GetProcesses();
            var subList = processlist.Where(x => x.ProcessName == "EXCEL").ToList();
            if (subList.Count > 0)
            {
                string excelListStr = "";
                foreach (var aExcelProcess in subList)
                {
                    excelListStr += $"Process id {aExcelProcess.Id} and name {aExcelProcess.ProcessName}\n\r";
                }
                MessageBox.Show(excelListStr + "is running");
                List<Workbook> listOfWorkbooks = Helper.ExcelGetRunningOjbects();

                MessageBox.Show($"Number of workbooks {listOfWorkbooks.Count}");
                foreach (var aWkb in listOfWorkbooks)
                {
                    string aPath = aWkb.Path;
                    aWkb.Application.Visible = true;
                }




            }
            else
            {
                MessageBox.Show($"No instances of Excel is running");
            }
        }

        private void btnSelectMainTool_Click(object sender, EventArgs e)
        {
            OpenFileDialog AppToolMain = new OpenFileDialog();
            AppToolMain.InitialDirectory = Path.GetDirectoryName(persistentData.AppToolMainPathFile);
            AppToolMain.RestoreDirectory = true;
            AppToolMain.Title = "Select main OE file";
            AppToolMain.DefaultExt = "*.xlsm";
            if (AppToolMain.ShowDialog() == DialogResult.OK)
            {
                lblOEdbMain.Text = AppToolMain.FileName;
                persistentData.AppToolMainPathFile = AppToolMain.FileName;
                persistentData.SavePersistentData(persistentData);


                //excelApp.Visible = true;

                //aWorkbook = excelApp.Workbooks.Open(LblSourceFilePlan2.Text.ToString());

            }
        }

        private void btnSelectImportFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog OEsource = new OpenFileDialog();
            OEsource.InitialDirectory = Path.GetDirectoryName(persistentData.AppToolSourcePathFile);
            OEsource.RestoreDirectory = true;
            OEsource.Title = "Select source OE file";
            OEsource.DefaultExt = "*.xlsm";
            if (OEsource.ShowDialog() == DialogResult.OK)
            {
                lblOEdbSource.Text = OEsource.FileName;
                persistentData.AppToolSourcePathFile = OEsource.FileName;
                persistentData.SavePersistentData(persistentData);


                //excelApp.Visible = true;

                //aWorkbook = excelApp.Workbooks.Open(LblSourceFilePlan2.Text.ToString());

            }
        }

        private void btnSetupMain2_Click(object sender, EventArgs e)
        {
            {
                // *** Setup the main sheet
                //excelApp.Visible = true;

                // *** Load main
                //mainWorkbook = excelApp.Workbooks.Open(persistentData.AppToolMainPathFile);

                if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(persistentData.AppToolSourcePathFile)) == null)
                {
                    //MessageBox.Show($"File {Path.GetFileName(persistentData.AppToolMainPathFile)}" +
                    //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.AppToolMainPathFile)}" +
                    //    "\n\rdoes not exists");
                    MessageBox.Show("File not found, has it been moved or deleted?");
                    return;
                }


                lblStatus.Text = "";
                string statusStr = "";
                foreach (Worksheet wksMain in mainWorkbook.Worksheets)
                {
                    switch (wksMain.Name)
                    {

                        case "CAR":
                        case "CM":
                        case "DAR":
                        case "EST":
                        case "MC":
                        case "MPM":
                        case "OT":
                        case "PAD":
                        case "PCM":
                        case "PLAN":
                        case "PQA":
                        case "PR":
                        case "RDM":
                        case "RSK":
                        case "VV":
                        case "PI":
                        case "TS":
                            // *** Find the number of rows
                            int NumberOfRows = Helper.FindEndOfWorksheet(wksMain, 2, 2, 50000);

                            // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                            Range mainRange = wksMain.Range["F" + 2, "Q" + NumberOfRows];
                            object[,] mainValue = mainRange.Value;

                            Range scopeRange = wksMain.Range["C" + 2, "C" + NumberOfRows];
                            object[,] scopeValue = scopeRange.Value;


                            // *** show progress
                            statusStr = statusStr + wksMain.Name + ".";
                            lblStatus.Text = statusStr;

                            // *** search rows for for upload
                            for (int rowS = 1; rowS <= (NumberOfRows - 2 + 1); rowS++)
                            {
                                if (scopeValue[rowS, 1] != null)
                                {
                                    if (scopeValue[rowS, 1].ToString() != "OoS") //|| scopeValue[rowS,1].ToString() == "DM")
                                    {
                                        // *** Item to initialise
                                        mainValue[rowS, 1] = 0;
                                        mainValue[rowS, 2] = 0;
                                        mainValue[rowS, 3] = 0;
                                        //mainRange[rowS, 4] = "Weakness";
                                        mainValue[rowS, 5] = 1;
                                        //mainRange[rowS, 6] = "Strength";
                                        mainValue[rowS, 7] = "OEdb PA:" + wksMain.Name;  // Data Collection Source OE Notes
                                        mainValue[rowS, 8] = selectInterviewAndName(wksMain.Name);// "[x] Interview - Name"; // Affirmation Source (session and person)
                                                                                                  //mainRange[rowS, 9] = "Questions";
                                                                                                  //mainRange[rowS, 10] = 1;
                                    }
                                }
                            }
                            mainRange.Value = mainValue;

                            break;


                        case "GOV":
                        case "II":
                            // *** Find the number of rows
                            int NumberOfRowsIIG;
                            if (wksMain.Name == "II") NumberOfRowsIIG = Helper.FindEndOfWorksheet(wksMain, 3, 2, cIIMaxRows);
                            else NumberOfRowsIIG = Helper.FindEndOfWorksheet(wksMain, 3, 2, cGOVMaxRows);

                            //int NumberOfRowsIIG = Helper.FindEndOfWorksheetBrute(wksMain, 3, 2, 4805);

                            // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                            Range mainRangeIIG = wksMain.Range["G" + 2, "R" + NumberOfRowsIIG];
                            object[,] mainValueIIG = mainRangeIIG.Value;

                            Range scopeRangeIIG = wksMain.Range["D" + 2, "D" + NumberOfRowsIIG];
                            object[,] scopeValueIIG = scopeRangeIIG.Value;


                            // *** show progress
                            statusStr = statusStr + wksMain.Name + ".";
                            lblStatus.Text = statusStr;

                            // *** search rows for for upload
                            for (int rowS = 1; rowS <= (NumberOfRowsIIG - 2 + 1); rowS++)
                            {
                                if (scopeValueIIG[rowS, 1] != null)
                                {

                                    if (scopeValueIIG[rowS, 1].ToString() != "OoS") // scopeValueIIG[rowS, 1].ToString() == "DM" || 
                                    {
                                        // *** Item to initialise
                                        mainValueIIG[rowS, 1] = 0;
                                        mainValueIIG[rowS, 2] = 0;
                                        mainValueIIG[rowS, 3] = 0;
                                        //mainRange[rowS, 4] = "Weakness";
                                        mainValueIIG[rowS, 5] = 1;
                                        //mainRange[rowS, 6] = "Strength";
                                        mainValueIIG[rowS, 7] = "OEdb PA:" + wksMain.Name; // Data Collection Source OE Notes
                                        mainValueIIG[rowS, 8] = "[x] Interview - Name"; // Affirmation Source (session and person)
                                                                                        //mainRange[rowS, 9] = "Questions";
                                                                                        //mainRange[rowS, 10] = 1;

                                    }
                                }

                            }
                            mainRangeIIG.Value = mainValueIIG;

                            break;


                    }

                }
                statusStr = statusStr + "done";
                lblStatus.Text = statusStr;

                MessageBox.Show("Done");

                // Iterate through source, if you find [upload], change to [done-upload] AND copy full line to main

            }
        }

        private void btnInsertInterviews2_Click(object sender, EventArgs e)
        {
            // *** Setup the main sheet
            //excelApp.Visible = true;

            // *** Load main
            //mainWorkbook = excelApp.Workbooks.Open(persistentData.AppToolMainPathFile);
            ;
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(persistentData.AppToolMainPathFile)) == null)
            {
                //MessageBox.Show($"File {Path.GetFileName(persistentData.AppToolMainPathFile)}" +
                //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.AppToolMainPathFile)}" +
                //    "\n\rdoes not exists");
                MessageBox.Show("File not found, has it been moved or deleted?");
                return;
            }

            lblStatus.Text = "";
            string statusStr = "";
            foreach (Worksheet wksMain in mainWorkbook.Worksheets)
            {
                switch (wksMain.Name)
                {

                    case "CAR":
                    case "CM":
                    case "DAR":
                    case "EST":
                    case "MC":
                    case "MPM":
                    case "OT":
                    case "PAD":
                    case "PCM":
                    case "PLAN":
                    case "PQA":
                    case "PR":
                    case "RDM":
                    case "RSK":
                    case "VV":
                    case "PI":
                    case "TS":
                        // *** Find the number of rows
                        int NumberOfRows = Helper.FindEndOfWorksheet(wksMain, 2, 2, 50000);

                        // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                        Range mainRange = wksMain.Range["F" + 2, "Q" + NumberOfRows];
                        object[,] mainValue = mainRange.Value;

                        Range scopeRange = wksMain.Range["C" + 2, "C" + NumberOfRows];
                        object[,] scopeValue = scopeRange.Value;


                        // *** show progress
                        statusStr = statusStr + wksMain.Name + ".";
                        lblStatus.Text = statusStr;

                        // *** search rows for for upload
                        for (int rowS = 1; rowS <= (NumberOfRows - 2 + 1); rowS++)
                        {
                            if (scopeValue[rowS, 1] != null)
                            {
                                if (scopeValue[rowS, 1].ToString() != "OoS") //|| scopeValue[rowS,1].ToString() == "DM")
                                {
                                    // *** Item to initialise
                                    // mainValue[rowS, 1] = 0;
                                    // mainValue[rowS, 2] = 0;
                                    // mainValue[rowS, 3] = 0;
                                    //mainRange[rowS, 4] = "Weakness";
                                    //   mainValue[rowS, 5] = 1;
                                    //mainRange[rowS, 6] = "Strength";
                                    //   mainValue[rowS, 7] = "OEdb PA:" + wksMain.Name;  // Data Collection Source OE Notes
                                    mainValue[rowS, 8] = selectInterviewAndName(wksMain.Name);// "[x] Interview - Name"; // Affirmation Source (session and person)
                                                                                              //mainRange[rowS, 9] = "Questions";
                                                                                              //mainRange[rowS, 10] = 1;
                                }
                            }
                        }
                        mainRange.Value = mainValue;

                        break;


                    case "GOV":
                    case "II":
                        // *** Find the number of rows
                        int NumberOfRowsIIG;
                        if (wksMain.Name == "II") NumberOfRowsIIG = Helper.FindEndOfWorksheet(wksMain, 3, 2, cIIMaxRows);
                        else NumberOfRowsIIG = Helper.FindEndOfWorksheet(wksMain, 3, 2, cGOVMaxRows);

                        // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                        Range mainRangeIIG = wksMain.Range["G" + 2, "R" + NumberOfRowsIIG];
                        object[,] mainValueIIG = mainRangeIIG.Value;

                        Range scopeRangeIIG = wksMain.Range["D" + 2, "D" + NumberOfRowsIIG];
                        object[,] scopeValueIIG = scopeRangeIIG.Value;


                        // *** show progress
                        statusStr = statusStr + wksMain.Name + ".";
                        lblStatus.Text = statusStr;

                        // *** search rows for for upload
                        for (int rowS = 1; rowS <= (NumberOfRowsIIG - 2 + 1); rowS++)
                        {
                            if (scopeValueIIG[rowS, 1] != null)
                            {

                                if (scopeValueIIG[rowS, 1].ToString() != "OoS") // scopeValueIIG[rowS, 1].ToString() == "DM" || 
                                {
                                    // *** Item to initialise
                                    //  mainValueIIG[rowS, 1] = 0;
                                    //  mainValueIIG[rowS, 2] = 0;
                                    //  mainValueIIG[rowS, 3] = 0;
                                    //mainRange[rowS, 4] = "Weakness";
                                    //   mainValueIIG[rowS, 5] = 1;
                                    //mainRange[rowS, 6] = "Strength";
                                    //   mainValueIIG[rowS, 7] = "OEdb PA:" + wksMain.Name; // Data Collection Source OE Notes
                                    mainValueIIG[rowS, 8] = "[x] Interview - Name"; // Affirmation Source (session and person)
                                                                                    //mainRange[rowS, 9] = "Questions";
                                                                                    //mainRange[rowS, 10] = 1;

                                }
                            }

                        }
                        mainRangeIIG.Value = mainValueIIG;

                        break;


                }

            }
            statusStr = statusStr + "done";
            lblStatus.Text = statusStr;

            MessageBox.Show("Done");
        }

        private void btnMergeSources2_Click(object sender, EventArgs e)
        {
            // *** Merge source inot main
            //excelApp.Visible = true;
            // excelApp.MacroOptions2(XlRunAutoMacro.xlAutoDeactivate );
            //= XlRunAutoMacro.xlAutoDeactivate;


            // *** Load source
            //sourceWorkbook = excelApp.Workbooks.Open(persistentData.AppToolSourcePathFile);
            if ((sourceWorkbook = Helper.CheckIfOpenAndOpenXlsx(persistentData.AppToolSourcePathFile)) == null)
            {
                //MessageBox.Show($"File {Path.GetFileName(persistentData.AppToolSourcePathFile)}" +
                //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.AppToolSourcePathFile)}" +
                //    "\n\rdoes not exists");
                MessageBox.Show("File not found, has it been moved or deleted?");
                return;
            }
            // *** Load main
            //mainWorkbook = excelApp.Workbooks.Open(persistentData.AppToolMainPathFile);
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(persistentData.AppToolMainPathFile)) == null)
            {
                //MessageBox.Show($"File {Path.GetFileName(persistentData.AppToolMainPathFile)}" +
                //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.AppToolMainPathFile)}" +
                //    "\n\rdoes not exists");
                MessageBox.Show("File not found, has it been moved or deleted?");
                return;
            }

            //  int row;
            //  string sValueN;
            Worksheet wsMain;

            lblStatus.Text = "";
            string statusStr = "";
            foreach (Worksheet wsSource in sourceWorkbook.Sheets)
            {

                switch (wsSource.Name)
                {

                    case "CAR":
                    case "CM":
                    case "DAR":
                    case "EST":
                    case "MC":
                    case "MPM":
                    case "OT":
                    case "PAD":
                    case "PCM":
                    case "PLAN":
                    case "PQA":
                    case "PR":
                    case "RDM":
                    case "RSK":
                    case "VV":
                    case "PI":
                    case "TS":
                        wsMain = mainWorkbook.Worksheets[wsSource.Name];
                        // *** Search string

                        // *** Find the number of rows
                        int NumberOfRows = Helper.FindEndOfWorksheet(wsSource, 2, 2, 50000);

                        // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                        Range soureRange = wsSource.Range["F" + 2, "Q" + NumberOfRows];
                        object[,] sourceValue = soureRange.Value;

                        Range mainRange = wsMain.Range["F" + 2, "Q" + NumberOfRows];
                        object[,] mainValue = mainRange.Value;

                        // *** show progress
                        statusStr = statusStr + wsSource.Name + ".";
                        lblStatus.Text = statusStr;

                        // *** search rows for for upload
                        for (int rowS = 1; rowS <= (NumberOfRows - 2 + 1); rowS++)
                        {
                            if (sourceValue[rowS, 9] != null)
                            {
                                string orriginal = sourceValue[rowS, 9].ToString();
                                string orriginalUpper = orriginal.ToUpper();

                                int first = orriginalUpper.IndexOf("[UPLOAD]");
                                if (first < 0) first = orriginalUpper.IndexOf("(UPLOAD)");

                                if (first >= 0)
                                { // Upload the row
                                    string part1 = orriginal.Substring(0, first);
                                    int part2Count = orriginal.Length - part1.Length - "[UPLOAD]".Length;
                                    string part2 = orriginal.Substring(first + "[UPLOAD]".Length, part2Count);
                                    sourceValue[rowS, 9] = part1 + " [*uploaded*] " + part2;//  "[*Upload done*]";
                                    for (int col = 1; col <= 12; col++)
                                    {
                                        mainValue[rowS, col] = sourceValue[rowS, col];
                                    }
                                    //                                    mainValue[rowS, 9] = "[*Uploaded*]";
                                }
                            }
                        }

                        mainRange.Value = mainValue;
                        soureRange.Value = sourceValue;


                        break;


                    case "GOV":
                    case "II":
                        wsMain = mainWorkbook.Worksheets[wsSource.Name];
                        // *** Find the number of rows

                        int NumberOfRowsIIG;
                        if (wsSource.Name == "II") NumberOfRowsIIG = Helper.FindEndOfWorksheet(wsSource, 3, 2, cIIMaxRows);
                        else NumberOfRowsIIG = Helper.FindEndOfWorksheet(wsSource, 3, 2, cGOVMaxRows);


                        //int NumberOfRowsIIG = Helper.FindEndOfWorksheetBrute(wsSource, 3, 2, 4805);


                        // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                        Range soureRangeIIG = wsSource.Range["G" + 2, "R" + NumberOfRowsIIG];
                        object[,] sourceValueIIG = soureRangeIIG.Value;

                        Range mainRangeIIG = wsMain.Range["G" + 2, "R" + NumberOfRowsIIG];
                        object[,] mainValueIIG = mainRangeIIG.Value;

                        // *** show progress
                        statusStr = statusStr + wsSource.Name + ".";
                        lblStatus.Text = statusStr;

                        // *** search rows for for upload
                        for (int rowS = 1; rowS <= (NumberOfRowsIIG - 2 + 1); rowS++)
                        {
                            if (sourceValueIIG[rowS, 9] != null)
                            {
                                string orriginal = sourceValueIIG[rowS, 9].ToString();
                                string orriginalUpper = orriginal.ToUpper();

                                int first = orriginalUpper.IndexOf("[UPLOAD]");
                                if (first < 0) first = orriginalUpper.IndexOf("(UPLOAD)");

                                if (first >= 0)
                                { // Upload the row
                                    string part1 = orriginal.Substring(0, first);
                                    int part2Count = orriginal.Length - part1.Length - "[UPLOAD]".Length;
                                    string part2 = orriginal.Substring(first + "[UPLOAD]".Length, part2Count);
                                    sourceValueIIG[rowS, 9] = part1 + " [*uploaded*] " + part2;//  "[*Upload done*]";
                                    for (int col = 1; col <= 12; col++)
                                    {
                                        mainValueIIG[rowS, col] = sourceValueIIG[rowS, col];
                                    }
                                    //                                    mainValue[rowS, 9] = "[*Uploaded*]";
                                }
                            }
                        }

                        // *** search rows for for upload
                        soureRangeIIG.Value = sourceValueIIG;
                        mainRangeIIG.Value = mainValueIIG;
                        break;
                }

            }
            statusStr = statusStr + "done";
            lblStatus.Text = statusStr;

            MessageBox.Show("Done");

            // Iterate through source, if you find [upload], change to [done-upload] AND copy full line to main


        }

        private void btnHideOoS2_Click(object sender, EventArgs e)
        {
            // *** Setup the main sheet
            //excelApp.Visible = true;

            // *** Load main
            // mainWorkbook = excelApp.Workbooks.Open(persistentData.AppToolMainPathFile);

            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(persistentData.AppToolMainPathFile)) == null)
            {
                //MessageBox.Show($"File {Path.GetFileName(persistentData.AppToolMainPathFile)}" +
                //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.AppToolMainPathFile)}" +
                //    "\n\rdoes not exists");
                MessageBox.Show("File not found, has it been moved or deleted?");
                return;
            }


            //excelApp.Application.WindowState = XlWindowState.xlMinimized;

            lblStatus.Text = "";
            string statusStr = "";
            foreach (Worksheet wksMain in mainWorkbook.Worksheets)
            {
                switch (wksMain.Name)
                {

                    case "CAR":
                    case "CM":
                    case "DAR":
                    case "EST":
                    case "MC":
                    case "MPM":
                    case "OT":
                    case "PAD":
                    case "PCM":
                    case "PLAN":
                    case "PQA":
                    case "PR":
                    case "RDM":
                    case "RSK":
                    case "VV":
                    case "PI":
                    case "TS":
                        HelperHideRows2(wksMain, 2, 2, 20000, 1, 3);
                        statusStr = statusStr + "." + wksMain.Name;
                        break;

                    case "GOV":
                        HelperHideRows2(wksMain, 3, 2, cGOVMaxRows, 1, 4);
                        statusStr = statusStr + "." + wksMain.Name;
                        break;

                    case "II":
                        HelperHideRows2(wksMain, 3, 2, cIIMaxRows, 2, 4);
                        statusStr = statusStr + "." + wksMain.Name;
                        break;


                }
                lblStatus.Text = statusStr;
            }
        }



        private void btnExtractFindings_Click(object sender, EventArgs e)
        {

            // *** Load main CMMI tool
            // mainWorkbook = excelApp.Workbooks.Open(persistentData.AppToolMainPathFile);
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(persistentData.AppToolMainPathFile)) == null)
            {
                //MessageBox.Show($"File {Path.GetFileName(persistentData.AppToolMainPathFile)}" +
                //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.AppToolMainPathFile)}" +
                //    "\n\rdoes not exists");
                MessageBox.Show("File not found, has it been moved or deleted?");
                return;
            }

            // *** Does the main workbook contain a findings sheet, if not add one, if it does, assign it and clear it
            Worksheet findingsWks = AssignOrCreateWorksheet(mainWorkbook, "Findings");
            findingsWks.Range["A:C"].Clear();
            findingsWks.Cells[1, 1].Value = "PA";
            findingsWks.Cells[1, 2].Value = "Strength/Weakness/Improvement";
            findingsWks.Cells[1, 3].Value = "Description";


            int findigsRow = 2;

            lblStatus.Text = "";
            string statusStr = "";
            foreach (Worksheet wksMain in mainWorkbook.Worksheets)
            {
                switch (wksMain.Name)
                {

                    case "CAR":
                    case "CM":
                    case "DAR":
                    case "EST":
                    case "MC":
                    case "MPM":
                    case "OT":
                    case "PAD":
                    case "PCM":
                    case "PLAN":
                    case "PQA":
                    case "PR":
                    case "RDM":
                    case "RSK":
                    case "VV":
                    case "PI":
                    case "TS":
                        HelperExtractFindings(wksMain, findingsWks, cXXSearchNumberOfWksRowsCol, cMostPAStartRow, cMostPAEndRow, cPAtestColumn,
                            cMostPAtestOoS, cXXWeaknessCol, cXXStrengthCol, cXXQuestionCol, cXXImprovementCol, ref findigsRow);
                        statusStr = statusStr + "." + wksMain.Name;
                        break;

                    case "GOV":
                        HelperExtractFindings(wksMain, findingsWks, cGOVandIIPASearchNumberOfWksRowsCol, cMostPAStartRow, cGOVMaxRows, cPAtestColumn,
                            cIIandGOVOoSTestCol, cIIandGOVWeaknessCol, cIIandGOVStrenghtCol, cIIandGOVQuestionCol, cIIandGOVImprovementCol, ref findigsRow);
                        statusStr = statusStr + "." + wksMain.Name;
                        break;

                    case "II":
                        HelperExtractFindings(wksMain, findingsWks, cGOVandIIPASearchNumberOfWksRowsCol, cMostPAStartRow, cIIMaxRows, cPAtestColumn,
                            cIIandGOVOoSTestCol, cIIandGOVWeaknessCol, cIIandGOVStrenghtCol, cIIandGOVQuestionCol, cIIandGOVImprovementCol, ref findigsRow);
                        statusStr = statusStr + "." + wksMain.Name;
                        break;


                }
                lblStatus.Text = statusStr;
            }
            //  wksMain.Application.Visible = true;
            findingsWks.Activate();
            MessageBox.Show("Findings extracted");
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="wksMain"></param>
        /// <param name="wksFindings"></param>
        /// <param name="searchForEndOfWksColumn">Used to find the end of the worksheet</param>
        /// <param name="startRow">Where to start the search</param>
        /// <param name="EndRow">Maximum rows expected for the sheet</param>
        /// <param name="testColumnPA">Where we expect the PA to be</param>
        /// <param name="testColumnOoS">Where we expect the OoS status to be</param>
        /// <param name="weaknessCol">Where we expect the weakness statements to be</param>
        /// <param name="strengthCol"></param>
        /// <param name="improvmentCol"></param>
        /// <param name="findigsRow"></param>
        const int CD_Heading = 1;
        const int CD_practiceCol = 2;
        const int CD_weaknessCol = 12;
        const int CD_strengthCol = 13;
        const int CD_recommendationCol = 14;

        private void HelperExtractFindingsDemixOE(Worksheet wksMain, Worksheet wksFindings, int searchForEndOfWksColumn, int startRow, int EndRow, ref int findigsRow)
        {
            //  wksMain.Application.Visible = false;

            // *** Find the number of rows
            int NumberOfRows = Helper.FindEndOfWorksheet(wksMain, searchForEndOfWksColumn, startRow, EndRow);

            // Range mainRange = wksMain.Range["A" + startRow, "Z" + NumberOfRows];
            // object[,] mainValue = mainRange.Value;


            // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
            //string practiceStr = "";
            // for (int rowS = startRow; rowS <= NumberOfRows; rowS++)
            for (int rowS = startRow; rowS <= EndRow; rowS++)
            {

                // *** test if it is of type "2 Prac_OU"
                if (wksMain.Cells[rowS, CD_Heading]?.Value == "2 Prac_OU")
                {
                    string practiceStr = wksMain.Cells[rowS, CD_practiceCol]?.Value?.ToString();

                    string weaknessStr2 = wksMain.Cells[rowS, CD_weaknessCol]?.Value?.ToString();
                    if (!string.IsNullOrEmpty(weaknessStr2))
                    {
                        wksFindings.Cells[findigsRow, 1].Value = practiceStr;
                        wksFindings.Cells[findigsRow, 2].Value = "Weakness";
                        wksFindings.Cells[findigsRow, 3].Value = weaknessStr2;
                        findigsRow++;
                    }

                    string strengthStr2 = wksMain.Cells[rowS, CD_strengthCol]?.Value?.ToString();
                    if (!string.IsNullOrEmpty(strengthStr2))
                    {
                        wksFindings.Cells[findigsRow, 1].Value = practiceStr;
                        wksFindings.Cells[findigsRow, 2].Value = "Strength";
                        wksFindings.Cells[findigsRow, 3].Value = strengthStr2;
                        findigsRow++;
                    }

                    string recommendationStr2 = wksMain.Cells[rowS, CD_recommendationCol]?.Value?.ToString();
                    if (!string.IsNullOrEmpty(recommendationStr2))
                    {
                        wksFindings.Cells[findigsRow, 1].Value = practiceStr;
                        wksFindings.Cells[findigsRow, 2].Value = "Recommendation";
                        wksFindings.Cells[findigsRow, 3].Value = recommendationStr2;
                        findigsRow++;
                    }

                }



            }


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wksMain"></param>
        /// <param name="wksFindings"></param>
        /// <param name="searchForEndOfWksColumn">Used to find the end of the worksheet</param>
        /// <param name="startRow">Where to start the search</param>
        /// <param name="EndRow">Maximum rows expected for the sheet</param>
        /// <param name="testColumnPA">Where we expect the PA to be</param>
        /// <param name="testColumnOoS">Where we expect the OoS status to be</param>
        /// <param name="weaknessCol">Where we expect the weakness statements to be</param>
        /// <param name="strengthCol"></param>
        /// <param name="improvmentCol"></param>
        /// <param name="findigsRow"></param>
        private void HelperExtractFindings(Worksheet wksMain, Worksheet wksFindings, int searchForEndOfWksColumn, int startRow, int EndRow,
            int testColumnPA,
            int testColumnOoS, int weaknessCol, int strengthCol, int questionCol, int improvmentCol, ref int findigsRow)
        {
            //  wksMain.Application.Visible = false;

            // *** Find the number of rows
            int NumberOfRows = Helper.FindEndOfWorksheet(wksMain, searchForEndOfWksColumn, startRow, EndRow);

            Range mainRange = wksMain.Range["A" + startRow, "Z" + NumberOfRows];
            object[,] mainValue = mainRange.Value;


            // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
            string practiceStr = "";
            // for (int rowS = startRow; rowS <= NumberOfRows; rowS++)
            for (int rowS = 1; rowS <= NumberOfRows - startRow; rowS++)
            {
                string practiceStrTest = mainValue[rowS, testColumnPA]?.ToString(); // wksMain.Cells[rowS, testColumnPA]?.Value?.ToString();
                if (!string.IsNullOrEmpty(practiceStrTest)) practiceStr = practiceStrTest; // Set the practiceStr to the latest found practice number

                bool hiddenRow = true; // Default hide it
                string cell1 = mainValue[rowS, testColumnPA]?.ToString(); //wksMain.Cells[rowS, testColumnPA]?.Value?.ToString();
                string cell2 = mainValue[rowS, testColumnOoS]?.ToString();  //wksMain.Cells[rowS, testColumnOoS]?.Value?.ToString();
                if (cell1 == null && cell2 != "OoS") hiddenRow = false; // Show if null and NOT OoS
                if (cell1 != null) hiddenRow = false; // Show if not null
                                                      // Range rowRange = wksMain.Rows[rowS]; // mainRange[rowS,1].EntireRow.Hidden = hiddenRow;
                if (!hiddenRow)
                {
                    // *** test for weakness string
                    string weaknessStr = mainValue[rowS, weaknessCol]?.ToString(); //wksMain.Cells[rowS, weaknessCol]?.Value?.ToString();
                    if (!string.IsNullOrEmpty(weaknessStr))
                    {
                        wksFindings.Cells[findigsRow, 1].Value = practiceStr;
                        wksFindings.Cells[findigsRow, 2].Value = "Weakness";
                        wksFindings.Cells[findigsRow, 3].Value = weaknessStr;
                        findigsRow++;
                    }

                    // *** test for streangt string
                    string strengthStr = mainValue[rowS, strengthCol]?.ToString();  // wksMain.Cells[rowS, strengthCol]?.Value?.ToString();
                    if (!string.IsNullOrEmpty(strengthStr))
                    {
                        wksFindings.Cells[findigsRow, 1].Value = practiceStr;
                        wksFindings.Cells[findigsRow, 2].Value = "Strength";
                        wksFindings.Cells[findigsRow, 3].Value = strengthStr;
                        findigsRow++;
                    }

                    // *** test for question string
                    string questionStr = mainValue[rowS, questionCol]?.ToString();  // wksMain.Cells[rowS, improvmentCol]?.Value?.ToString();
                    if (!string.IsNullOrEmpty(questionStr))
                    {
                        wksFindings.Cells[findigsRow, 1].Value = practiceStr;
                        wksFindings.Cells[findigsRow, 2].Value = "Question";
                        wksFindings.Cells[findigsRow, 3].Value = questionStr;
                        findigsRow++;
                    }

                    // *** test for improvement string
                    string improvementStr = mainValue[rowS, improvmentCol]?.ToString();  // wksMain.Cells[rowS, improvmentCol]?.Value?.ToString();
                    if (!string.IsNullOrEmpty(improvementStr))
                    {
                        wksFindings.Cells[findigsRow, 1].Value = practiceStr;
                        wksFindings.Cells[findigsRow, 2].Value = "Improvement";
                        wksFindings.Cells[findigsRow, 3].Value = improvementStr;
                        findigsRow++;
                    }


                }
            }


        }

        private Worksheet AssignOrCreateWorksheet(Workbook aWkb, string wksName)
        {
            throw new ApplicationException("This method should not be called anymore from here!");
            foreach (Worksheet aWks in aWkb.Worksheets)
            {
                if (aWks.Name.ToUpper() == wksName.ToUpper())
                {
                    return aWks;
                }
            }
            // No wks found
            Worksheet newWks = aWkb.Worksheets.Add(After: aWkb.Worksheets["Processes"]);
            newWks.Name = wksName;
            return newWks;
        }

        private void btnSelectOEdb2_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog_OEdb = new OpenFileDialog();
            fileDialog_OEdb.InitialDirectory = Path.GetDirectoryName(persistentData.OEdatabasePathFile);
            fileDialog_OEdb.RestoreDirectory = true;
            fileDialog_OEdb.Title = "Select OE database";
            fileDialog_OEdb.DefaultExt = "*.xlsm";
            if (fileDialog_OEdb.ShowDialog() == DialogResult.OK)
            {
                lblOEdbPathFile.Text = fileDialog_OEdb.FileName;
                persistentData.OEdatabasePathFile = fileDialog_OEdb.FileName;
                persistentData.SavePersistentData(persistentData);

                //excelApp.Visible = true;

                //aWorkbook = excelApp.Workbooks.Open(LblSourceFilePlan2.Text.ToString());

            }
        }

        private void btnTestAndEngl2_Click(object sender, EventArgs e)
        {
            // *** Setup the main sheet
            // excelApp.Visible = true;

            // *** Load main
            //mainWorkbook = excelApp.Workbooks.Open(persistentData.OEdatabasePathFile);

            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(persistentData.OEdatabasePathFile)) == null)
            {
                //MessageBox.Show($"File {Path.GetFileName(persistentData.OEdatabasePathFile)}" +
                //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.OEdatabasePathFile)}" +
                //    "\n\rdoes not exists");
                MessageBox.Show("File not found, has it been moved or deleted?");
                return;
            }
            string basePath = Path.GetDirectoryName(persistentData.OEdatabasePathFile);

            lblStatus.Text = "OEdb:";
            string statusStr = "";
            foreach (Worksheet wksOEdb in mainWorkbook.Worksheets)
            {
                switch (wksOEdb.Name)
                {

                    case "CAR":
                    case "CM":
                    case "DAR":
                    case "EST":
                    case "MC":
                    case "MPM":
                    case "OT":
                    case "PAD":
                    case "PCM":
                    case "PLAN":
                    case "PQA":
                    case "PR":
                    case "RDM":
                    case "RSK":
                    case "VV":
                    case "PI":
                    case "TS":
                    case "GOV":
                    case "II":

                        //if (wksOEdb.Name=="PI")
                        //{
                        //    int stop = 1;
                        //}
                        // *** Find the number of rows
                        int NumberOfRows = Helper.FindEndOfWorksheetBrute(wksOEdb, cOEnonEmptyColumn, cOEDatabaseHeadingStartRow, cOEDatabaseMaxRows);
                        Range columnToClear = wksOEdb.Range["Y:Z"];
                        columnToClear.Clear();

                        // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                        Range mainRange = wksOEdb.Range["A" + cOEDatabaseHeadingStartRow, "Z" + NumberOfRows];

                        // *** List all the hyperlinks https://www.e-iceblue.com/Tutorials/Spire.XLS/Spire.XLS-Program-Guide/Link/Retrieve-Hyperlinks-from-an-Excel-Sheet-in-C-VB.NET.html
                        Hyperlinks hyperLinkList = mainRange.Hyperlinks;
                        List<Hyperlink> hyperLinksToAdd = new List<Hyperlink>();

                        int hyperLinkRow;
                        int hyperLinkCol;
                        string hyperlinkAddress;
                        string PathFileToTest;
                        string PathEnglish;

                        string englishFullPathFile;
                        Boolean fileFound;

                        foreach (Hyperlink aHyperlink in hyperLinkList)
                        {
                            // *** Take each hyperlink and test it
                            hyperLinkRow = aHyperlink.Range.Row;
                            if (hyperLinkRow == 9 && wksOEdb.Name == "PI")
                            {
                                int stop = 1;
                            }
                            hyperLinkCol = aHyperlink.Range.Column;
                            hyperlinkAddress = aHyperlink.Address;

                            // *** Test if the file exists
                            fileFound = false;
                            PathFileToTest = Path.Combine(basePath, hyperlinkAddress);
                            if (File.Exists(PathFileToTest))
                            {
                                mainRange[hyperLinkRow - cOEDatabaseHeadingStartRow + 1, "y"].Value = "ok file";
                                fileFound = true;
                            }
                            else
                            {
                                if (Directory.Exists(PathFileToTest))
                                {
                                    mainRange[hyperLinkRow - cOEDatabaseHeadingStartRow + 1, "y"].Value = "ok directory";
                                }
                                else
                                {
                                    mainRange[hyperLinkRow - cOEDatabaseHeadingStartRow + 1, "y"].Value = "Not ok";
                                }
                            }

                            // *** Test if the english version exists
                            PathEnglish = Path.Combine(Path.GetDirectoryName(PathFileToTest), Path.GetFileNameWithoutExtension(PathFileToTest));
                            englishFullPathFile = "";
                            string theExtension = Path.GetExtension(PathFileToTest);
                            switch (theExtension.ToLower().Trim())
                            {
                                case ".xls":
                                case ".xlsx":
                                case ".xlsm":
                                    englishFullPathFile = PathEnglish + "-engl.xlsx";
                                    break;
                                case ".doc":
                                case ".docx":
                                case ".docm":
                                    englishFullPathFile = PathEnglish + "-engl.docx";
                                    break;
                                case ".ppt":
                                case ".pptx":
                                case ".pptm":
                                    englishFullPathFile = PathEnglish + "-engl.pptx";
                                    break;

                            }
                            if (englishFullPathFile != "")
                            {
                                // *** list the new hyperlink
                                if (fileFound && File.Exists(englishFullPathFile))
                                { // file exists, add it
                                  // mainRange[hyperLinkRow - cOEDatabaseHeadingStartRow + 1, "z"].Value = "engl";
                                    string remainderPath = englishFullPathFile.Remove(0, basePath.Length + 1);
                                    string formulaStr = "=hyperlink(\"" + remainderPath + "\",\"engl\")";
                                    mainRange[hyperLinkRow - cOEDatabaseHeadingStartRow + 1, "z"].Formula = formulaStr;

                                }
                                else
                                {
                                    mainRange[hyperLinkRow - cOEDatabaseHeadingStartRow + 1, "z"].Value = "none";
                                }
                            }
                        }

                        // *** Show the status
                        statusStr = statusStr + wksOEdb.Name + ".";
                        lblStatus.Text = statusStr;
                        break;
                }
            }
            statusStr = statusStr + "done";
            lblStatus.Text = statusStr;

            MessageBox.Show("Done");
        }

        private void btnIIGOVrating_Click(object sender, EventArgs e)
        {
            // *** Setup the main sheet
            //excelApp.Visible = true;

            // *** Load main
            // mainWorkbook = excelApp.Workbooks.Open(persistentData.AppToolMainPathFile);

            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(persistentData.AppToolMainPathFile)) == null)
            {
                //MessageBox.Show($"File {Path.GetFileName(persistentData.AppToolMainPathFile)}" +
                //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.AppToolMainPathFile)}" +
                //    "\n\rdoes not exists");
                MessageBox.Show("File not found, has it been moved or deleted?");
                return;
            }


            //excelApp.Application.WindowState = XlWindowState.xlMinimized;

            lblStatus.Text = "";
            string statusStr = "";

            Worksheet iigovCharWks = mainWorkbook.Worksheets["SP&OSF_rating"];
            iigovCharWks.Range["A4:Z200"].Clear();

            //    int SPOSFratingRow = 4;

            foreach (Worksheet wksMain in mainWorkbook.Worksheets)
            {
                switch (wksMain.Name)
                {

                    case "GOV":
                        HelperIIGOVcharacterization(wksMain, iigovCharWks, 3, 2, cGOVMaxRows, 1, 4);
                        //SP & OSF_rating(wksMain, 3, 2, cGOVMaxRows, 1, 4);
                        statusStr = statusStr + "." + wksMain.Name;
                        break;

                    case "II":
                        HelperIIGOVcharacterization(wksMain, iigovCharWks, 3, 2, cIIMaxRows, 1, 4);
                        //HelperHideRows2(wksMain, 3, 2, cIIMaxRows, 2, 4);
                        statusStr = statusStr + "." + wksMain.Name;
                        break;


                }
                lblStatus.Text = statusStr;
            }
        }

        private int ColumnOfPractice(string paStr)
        {
            int result = 17;
            switch (paStr.ToUpper().Trim())
            {
                case "II 1.1":
                    result = 3;
                    break;
                case "II 2.1":
                    result = 4;
                    break;
                case "II 2.2":
                    result = 5;
                    break;
                case "II 3.1":
                    result = 6;
                    break;
                case "II 3.2":
                    result = 7;
                    break;
                case "II 3.3":
                    result = 8;
                    break;
                case "GOV 1.1":
                    result = 9;
                    break;
                case "GOV 2.1":
                    result = 10;
                    break;
                case "GOV 2.2":
                    result = 11;
                    break;
                case "GOV 2.3":
                    result = 12;
                    break;
                case "GOV 2.4":
                    result = 13;
                    break;
                case "GOV 3.1":
                    result = 14;
                    break;
                case "GOV 3.2":
                    result = 15;
                    break;
                case "GOV 4.1":
                    result = 16;
                    break;
            }

            return result;
        }

        private void btnSelectWorkingDir_Click(object sender, EventArgs e)
        {
            // https://stackoverflow.com/questions/11624298/how-to-use-openfiledialog-to-select-a-folder
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = persistentData.LastAppraisalDirectory;
                DialogResult result = fbd.ShowDialog();



                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    persistentData.LastAppraisalDirectory = fbd.SelectedPath;
                    lblWorkingDirectory.Text = fbd.SelectedPath;

                    //// https://stackoverflow.com/questions/3218910/rename-a-file-in-c-sharp

                    //string[] files = Directory.GetFiles(fbd.SelectedPath);
                    //foreach (string fromFileName in files)
                    //{
                    //    // if you find a key, replace it and rename it
                    //    string toFileName = rgx.Replace(fromFileName, toStr);
                    //    if (string.Compare(fromFileName,toFileName) != 0)
                    //    {
                    //        MessageBox.Show($"From: {fromFileName} To: {toFileName}");
                    //    }
                    //}
                    // System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
                }
            }
        }

        private void btnChange_Click(object sender, EventArgs e)
        {

            // https://stackoverflow.com/questions/3218910/rename-a-file-in-c-sharp

            string[] files = Directory.GetFiles(lblWorkingDirectory.Text.ToString());

            Regex rgx = new Regex(txtFrom.Text.ToString());
            string toStr = txtTo.Text.ToString();

            int changedFiles = 0;
            foreach (string fromPathFileName in files)
            {
                // if you find a key, replace it and rename it
                string directoryPath = Path.GetDirectoryName(fromPathFileName);
                string fileName = Path.GetFileName(fromPathFileName);

                string toFileName = rgx.Replace(fileName, toStr);
                if (string.Compare(fileName, toFileName) != 0)
                {
                    // MessageBox.Show($"From: {fileName} To: {toFileName}");
                    try
                    {

                        System.IO.File.Move(Path.Combine(directoryPath, fileName), Path.Combine(directoryPath, toFileName));
                        changedFiles++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Cannot change file name error {ex.Message}");
                    }
                }
            }
            MessageBox.Show($"File names changed successfully={changedFiles}");
        }

        private void txtFrom_Leave(object sender, EventArgs e)
        {
            persistentData.FromText = txtFrom.Text.ToString();
        }

        private void txtTo_Leave(object sender, EventArgs e)
        {
            persistentData.ToText = txtTo.Text.ToString();
        }

        private void btnResetV01_Click(object sender, EventArgs e)
        {
            // https://stackoverflow.com/questions/3218910/rename-a-file-in-c-sharp

            string[] files = Directory.GetFiles(lblWorkingDirectory.Text.ToString());

            Regex rgx = new Regex(@"(v[0-9]+)|(V[0-9]+)");
            string toStr = "v01";

            int changedFiles = 0;
            foreach (string fromPathFileName in files)
            {
                // if you find a key, replace it and rename it
                string directoryPath = Path.GetDirectoryName(fromPathFileName);
                string fileName = Path.GetFileName(fromPathFileName);

                string toFileName = rgx.Replace(fileName, toStr);
                if (string.Compare(fileName, toFileName) != 0)
                {
                    // MessageBox.Show($"From: {fileName} To: {toFileName}");
                    try
                    {

                        System.IO.File.Move(Path.Combine(directoryPath, fileName), Path.Combine(directoryPath, toFileName));
                        changedFiles++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Cannot change file name error {ex.Message}");
                    }
                }
            }
            MessageBox.Show($"File names changed successfully={changedFiles}");
        }

        private void btnRemoveUpload_Click(object sender, EventArgs e)
        {

            // *** Load main
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(persistentData.AppToolMainPathFile)) == null)
            {
                MessageBox.Show("File not found, has it been moved or deleted?");
                return;
            }
            //          Worksheet wsMain;
            lblStatus.Text = "";
            string statusStr = "";
            foreach (Worksheet wsSource in mainWorkbook.Sheets)
            {

                switch (wsSource.Name)
                {

                    case "CAR":
                    case "CM":
                    case "DAR":
                    case "EST":
                    case "MC":
                    case "MPM":
                    case "OT":
                    case "PAD":
                    case "PCM":
                    case "PLAN":
                    case "PQA":
                    case "PR":
                    case "RDM":
                    case "RSK":
                    case "VV":
                    case "PI":
                    case "TS":
                        // *** Find the number of rows
                        int NumberOfRows = Helper.FindEndOfWorksheet(wsSource, 2, 2, 50000);

                        // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                        Range soureRange = wsSource.Range["F" + 2, "Q" + NumberOfRows];
                        object[,] sourceValue = soureRange.Value;

                        // *** show progress
                        statusStr = statusStr + wsSource.Name + ".";
                        lblStatus.Text = statusStr;

                        // *** search rows for for upload
                        for (int rowS = 1; rowS <= (NumberOfRows - 2 + 1); rowS++)
                        {
                            if (sourceValue[rowS, 9] != null)
                            {
                                string orriginal = sourceValue[rowS, 9].ToString();
                                string orriginalUpper = orriginal.ToUpper();

                                int first = orriginalUpper.IndexOf("[*UPLOADED*]");

                                if (first >= 0)
                                { // Remove the string
                                    string part1 = orriginal.Substring(0, first);
                                    int part2Count = orriginal.Length - part1.Length - "[*UPLOADED*]".Length;
                                    string part2 = orriginal.Substring(first + "[*UPLOADED*]".Length, part2Count);
                                    sourceValue[rowS, 9] = part1 + part2; // part1 + " [*uploaded*] " + part2;//  "[*Upload done*]";
                                }
                            }
                            // Clean all spaces
                            if (sourceValue[rowS, 4] != null) sourceValue[rowS, 4] = sourceValue[rowS, 4].ToString().Trim();
                            if (sourceValue[rowS, 6] != null) sourceValue[rowS, 6] = sourceValue[rowS, 6].ToString().Trim();
                            if (sourceValue[rowS, 9] != null) sourceValue[rowS, 9] = sourceValue[rowS, 9].ToString().Trim();
                            if (sourceValue[rowS, 12] != null) sourceValue[rowS, 12] = sourceValue[rowS, 12].ToString().Trim();
                        }

                        soureRange.Value = sourceValue;
                        break;
                    case "GOV":
                    case "II":
                        // *** Find the number of rows
                        int NumberOfRowsIIG;
                        if (wsSource.Name == "GOV") NumberOfRowsIIG = Helper.FindEndOfWorksheet(wsSource, 3, 2, cGOVMaxRows);
                        else NumberOfRowsIIG = Helper.FindEndOfWorksheet(wsSource, 3, 2, cIIMaxRows);

                        // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                        Range soureRangeIIG = wsSource.Range["G" + 2, "R" + NumberOfRowsIIG];
                        object[,] sourceValueIIG = soureRangeIIG.Value;

                        // *** show progress
                        statusStr = statusStr + wsSource.Name + ".";
                        lblStatus.Text = statusStr;

                        // *** search rows for for upload
                        for (int rowS = 1; rowS <= (NumberOfRowsIIG - 2 + 1); rowS++)
                        {
                            if (sourceValueIIG[rowS, 9] != null)
                            {
                                string orriginal = sourceValueIIG[rowS, 9].ToString();
                                string orriginalUpper = orriginal.ToUpper();

                                int first = orriginalUpper.IndexOf("[*UPLOADED*]");

                                if (first >= 0)
                                { // Remove the string
                                    string part1 = orriginal.Substring(0, first);
                                    int part2Count = orriginal.Length - part1.Length - "[*UPLOADED*]".Length;
                                    string part2 = orriginal.Substring(first + "[*UPLOADED*]".Length, part2Count);
                                    sourceValueIIG[rowS, 9] = (part1 + part2).Trim(); // part1 + " [*uploaded*] " + part2;//  "[*Upload done*]";
                                }
                            }
                            // Clean all spaces
                            if (sourceValueIIG[rowS, 4] != null) sourceValueIIG[rowS, 4] = sourceValueIIG[rowS, 4].ToString().Trim();
                            if (sourceValueIIG[rowS, 6] != null) sourceValueIIG[rowS, 6] = sourceValueIIG[rowS, 6].ToString().Trim();
                            if (sourceValueIIG[rowS, 9] != null) sourceValueIIG[rowS, 9] = sourceValueIIG[rowS, 9].ToString().Trim();
                            if (sourceValueIIG[rowS, 12] != null) sourceValueIIG[rowS, 12] = sourceValueIIG[rowS, 12].ToString().Trim();
                        }

                        soureRangeIIG.Value = sourceValueIIG;
                        break;
                }

            }
            statusStr = statusStr + "done";
            lblStatus.Text = statusStr;

            MessageBox.Show("Done");

            // Iterate through source, if you find [upload], change to [done-upload] AND copy full line to main

        }

        private void btnSelectQuestionFile_Click(object sender, EventArgs e)
        {
            // Remove from release 3.8.0.44


            ////OpenFileDialog questionsFileDialog = new OpenFileDialog();

            ////questionsFileDialog.InitialDirectory = Path.GetDirectoryName(persistentData.QuestionPathFile ?? @"c:\"); // cPath_start;
            ////questionsFileDialog.RestoreDirectory = true;
            ////questionsFileDialog.Title = "Select source file";
            ////questionsFileDialog.DefaultExt = "*.xlsx";
            ////if (questionsFileDialog.ShowDialog() == DialogResult.OK)
            ////{
            ////    lblQuestions.Text = questionsFileDialog.FileName;
            ////    // *** save new file
            ////    //persistentData.LastAppraisalDirectory = Path.GetDirectoryName(questionsFile.FileName);
            ////    //persistentData.CASPlanName = Path.GetFileName(questionsFile.FileName);
            ////    persistentData.QuestionPathFile = questionsFileDialog.FileName;
            ////    persistentData.SavePersistentData(persistentData);
            ////    lblQuestions.Text = persistentData.QuestionPathFile;
            ////    //lblWorkingDirectory.Text = persistentData.LastAppraisalDirectory;
            ////    //lblPlanName.Text = persistentData.CASPlanName;

            ////    //excelApp.Visible = true;

            ////    //aWorkbook = excelApp.Workbooks.Open(LblSourceFilePlan2.Text.ToString());

            ////    //if ((aWorkbook = Helper.CheckIfOpenAndOpen(persistentData.QuestionPathFile)) == null)
            ////    //{
            ////    //    //MessageBox.Show($"File {Path.GetFileName(LblSourceFilePlan2.Text.ToString())}" +
            ////    //    //    $"\n\rDirectory {Path.GetDirectoryName(LblSourceFilePlan2.Text.ToString())}" +
            ////    //    //    "\n\rdoes not exists");
            ////    //    MessageBox.Show("File not found, has it been moved or deleted?");
            ////    //    return;
            ////    //}

            ////    // Step 0: Clear the list to start afresh

            ////    //CMMIModel.Clear();

            ////}
        }

        private void btnImportModel_Click(object sender, EventArgs e)
        {
            // Remove from release 3.8.0.44


            ////// *** Test if the question file exists
            ////if (!File.Exists(persistentData.QuestionPathFile))
            ////{
            ////    MessageBox.Show($"The question file {persistentData.QuestionPathFile}\ndoes not exists!");
            ////    return;
            ////}
            ////else
            ////{
            ////    if ((questionWorkbook = Helper.CheckIfOpenAndOpen(persistentData.QuestionPathFile)) == null)
            ////    {
            ////        //MessageBox.Show($"File {Path.GetFileName(LblSourceFilePlan2.Text.ToString())}" +
            ////        //    $"\n\rDirectory {Path.GetDirectoryName(LblSourceFilePlan2.Text.ToString())}" +
            ////        //    "\n\rdoes not exists");
            ////        MessageBox.Show("File not found, has it been moved or deleted?");
            ////        return;
            ////    }

            ////    // Clear the model and start processin the questionWorkbook
            ////    CMMIModel.Clear();

            ////    //MessageBox.Show($"The question file exists, now processing it ... ");
            ////    string statusStr = "";
            ////    foreach (var worksheetName in Enum.GetValues(typeof(EPAcode))) //.Cast<EPAcode>().ToList())
            ////    {
            ////        // Open the worksheet and process it
            ////        PracticeArea aPracticeArea = Helper.ProcessPracticeArea(questionWorkbook, worksheetName.ToString());
            ////        if (aPracticeArea != null) CMMIModel.Add(aPracticeArea);

            ////        // Update status string
            ////        statusStr += worksheetName + " ";
            ////        lblStatus.Text = statusStr;
            ////    }
            ////    return;

            ////}



        }

        private void btnSelectDemixTool_Click(object sender, EventArgs e)
        {

            // Remove from release 3.8.0.44

            ////OpenFileDialog demixToolFileDialog = new OpenFileDialog();

            ////demixToolFileDialog.InitialDirectory = Path.GetDirectoryName(persistentData.DemixToolPathFile ?? @"c:\"); // cPath_start;
            ////demixToolFileDialog.RestoreDirectory = true;
            ////demixToolFileDialog.Title = "Select source file";
            ////demixToolFileDialog.DefaultExt = "*.xlsx";
            ////if (demixToolFileDialog.ShowDialog() == DialogResult.OK)
            ////{
            ////    lblDemixTool.Text = demixToolFileDialog.FileName;
            ////    persistentData.DemixToolPathFile = demixToolFileDialog.FileName;
            ////    persistentData.SavePersistentData(persistentData);
            ////    lblQuestions.Text = persistentData.DemixToolPathFile;
            ////}


        }

        private void btnGenerateFullTool_Click(object sender,
            EventArgs e)
        {
            // Remove from release 3.8.0.44
            ////DialogResult dialogResult = MessageBox.Show("Make sure Processess are correcly listed in tab:Project&Support! Continue?", "Warning", MessageBoxButtons.YesNo);
            ////if (dialogResult == DialogResult.Yes)
            ////{
            ////    //do something
            ////}
            ////else if (dialogResult == DialogResult.No)
            ////{
            ////    //do something else
            ////    return;
            ////}

            ////string[] mostPAs = { "PI", "TS", "PQA", "PR", "RDM", "VV", "MPM", "PAD", "PCM", "RSK", "OT", "EST", "MC", "PLAN", "CAR", "CM", "DAR", "SAM" };
            ////string[] specialPAs = { "GOV", "II" };
            ////const int cTemplateLevelRow = 3;
            ////const int cTemplateBlueRow = 4;
            ////const int cTemplateProcessRow = 5;
            ////const int cTemplateYellowRow = 6;
            ////const int cOERow = 7;

            ////// open demix tool, if not open
            ////Workbook demixToolWkb;
            ////if ((demixToolWkb = Helper.CheckIfOpenAndOpen(persistentData.DemixToolPathFile)) == null)
            ////{
            ////    MessageBox.Show("Cannot open the demix tool, is the file moved or deleted?");
            ////    return;
            ////}

            ////// generic variables
            ////Worksheet tmpl1Wks = demixToolWkb.Worksheets["Template1"];

            ////// demixToolWkb contains the opened workbook
            ////foreach (PracticeArea aPracticeArea in CMMIModel)
            ////{
            ////    // DEBUG CODE, SKIP most PAs
            ////    //if (mostPAs.Contains(aPracticeArea.PAcode.ToString())) continue;
            ////    // create a worksheet if it does not exist
            ////    //Worksheet aWks = Helper.OpenOrElseCreateWks(demixToolWkb, aPracticeArea.PAcode.ToString());
            ////    foreach (Worksheet findWks in demixToolWkb.Worksheets)
            ////    {
            ////        if (findWks.Name == aPracticeArea.PAcode.ToString()) findWks.Delete();
            ////    }
            ////    // Copy the template2 over that worksheet
            ////    Worksheet sourceWks;
            ////    Worksheet aWks;
            ////    sourceWks = demixToolWkb.Worksheets["Template2"];
            ////    //aWks = demixToolWkb.Worksheets.Add();
            ////    int numberOfWks = demixToolWkb.Worksheets.Count;
            ////    sourceWks.Copy(After: demixToolWkb.Worksheets[numberOfWks]);
            ////    aWks = demixToolWkb.Worksheets[numberOfWks + 1];
            ////    aWks.Name = aPracticeArea.PAcode.ToString();

            ////    // Setup the headings
            ////    aWks.Cells[1, 1].Value = aPracticeArea.Name;
            ////    aWks.Cells[2, 1].Value = aPracticeArea.NameChinese;
            ////    aWks.Cells[3, 2].Value = aPracticeArea.Intent;
            ////    aWks.Cells[4, 2].Value = aPracticeArea.IntentChinese;
            ////    aWks.Cells[5, 2].Value = aPracticeArea.Value;
            ////    aWks.Cells[6, 2].Value = aPracticeArea.ValueChinese;

            ////    // Setup parameters
            ////    int rowX = 9; // the starting row to process
            ////                  // Build each of the levels 
            ////    for (int levelX = 1; levelX <= 5; levelX++)
            ////    {

            ////        // Find all practices at this level
            ////        var levelPractices =
            ////            from aPractice in aPracticeArea.Practices
            ////            where aPractice.Level == levelX
            ////            orderby aPractice.Number
            ////            select aPractice;

            ////        if (levelPractices?.Count() >= 1)
            ////        {
            ////            // Practices found for this level
            ////            // Copy the level
            ////            Range levelRow = tmpl1Wks.Rows[cTemplateLevelRow];
            ////            Range destLevelRow = aWks.Rows[rowX];
            ////            levelRow.Copy(destLevelRow);

            ////            // Set the level number
            ////            aWks.Cells[rowX, 2].Value = $"Level {levelX}";
            ////            rowX++;

            ////            // run through each practice and process it
            ////            foreach (Practice aPractice in levelPractices)
            ////            {
            ////                // Copy the practice heading
            ////                Range blueRow = tmpl1Wks.Rows[cTemplateBlueRow];
            ////                Range destBlueRow = aWks.Rows[rowX];
            ////                blueRow.Copy(destBlueRow);
            ////                aWks.Cells[rowX, 2].Value = $"{aPractice.Acronym} {aPractice.Level}.{aPractice.Number}";
            ////                // Extract statement, work products, activities and questions
            ////                string statementStr, workProductStr, activityStr, questionStr;
            ////                Helper.ExtractPracticeAreaInformation(aPractice, out statementStr, out workProductStr,
            ////                    out activityStr, out questionStr);
            ////                aWks.Cells[rowX, 3].Value = statementStr;
            ////                aWks.Cells[rowX, 9].Value = workProductStr;
            ////                aWks.Cells[rowX, 10].Value = activityStr;
            ////                aWks.Cells[rowX, 11].Value = questionStr;

            ////                rowX++;

            ////                if (mostPAs.Contains(aPracticeArea.PAcode.ToString()))
            ////                {
            ////                    // process most PAs
            ////                    // Find all projects / support funcitons that has this practice sampled
            ////                    List<WorkUnit> workUnitsInScope = new List<WorkUnit>();
            ////                    foreach (WorkUnit aWorkUnit in WorkUnitList)
            ////                    {
            ////                        // See if the practice is present in the work unit practice list
            ////                        var matchingPAList = from aPAitem in aWorkUnit.PAlist
            ////                                             where aPAitem.PAcode == aPracticeArea.PAcode
            ////                                             select aPAitem;
            ////                        // If it is present, add it to the list
            ////                        if (matchingPAList?.Count() > 0)
            ////                        {
            ////                            workUnitsInScope.Add(aWorkUnit);
            ////                        }
            ////                    }

            ////                    // workUnitsInScope contains all the work units, so now add them to the sheet
            ////                    foreach (WorkUnit workUnitToAdd in workUnitsInScope)
            ////                    {
            ////                        // List the work unit in scope
            ////                        Range yelloRow = tmpl1Wks.Rows[cTemplateYellowRow];
            ////                        Range destYellowRow = aWks.Rows[rowX];
            ////                        yelloRow.Copy(destYellowRow);
            ////                        aWks.Cells[rowX, 2].Value = $"{aPractice.Acronym} {aPractice.Level}.{aPractice.Number}";
            ////                        aWks.Cells[rowX, 3].Value = workUnitToAdd.Name;

            ////                        // identify the interviewees
            ////                        List<Schedule2> scheduleForWorkUnit = Schedule2List.Where(x => x.PA == aPracticeArea.PAcode.ToString() && x.WorkID == workUnitToAdd.ID).ToList();
            ////                        if (scheduleForWorkUnit.Count > 0)
            ////                        {
            ////                            string meetingParticipantStr = "";
            ////                            bool firstReview = true;
            ////                            foreach (var aScheduleItem in scheduleForWorkUnit)
            ////                            {
            ////                                if (firstReview)
            ////                                {
            ////                                    meetingParticipantStr = $"{aScheduleItem.InterviewSession}({aScheduleItem.ParticipantName})";
            ////                                }
            ////                                else
            ////                                {
            ////                                    meetingParticipantStr = meetingParticipantStr + $" {aScheduleItem.InterviewSession}({aScheduleItem.ParticipantName})";
            ////                                }
            ////                            }
            ////                            aWks.Cells[rowX, 8].Value = meetingParticipantStr;
            ////                        }
            ////                        // List staff applicable to this project
            ////                        // var staffForThisWorkUnit = StaffList.Where(x => x.WorkID == workUnitToAdd.ID).ToList();

            ////                        //var listOfInterviewees = StaffList.Where(x => x.)
            ////                        rowX++;

            ////                        Range oeRow = tmpl1Wks.Rows[cOERow];
            ////                        for (int y = 0; y < 2; y++)
            ////                        {
            ////                            Range destOERow = aWks.Rows[rowX];
            ////                            oeRow.Copy(destOERow);
            ////                            aWks.Cells[rowX, 2].Value = workUnitToAdd.Name;
            ////                            rowX++;
            ////                        }
            ////                    }

            ////                }
            ////                else
            ////                {
            ////                    if (specialPAs.Contains(aPracticeArea.PAcode.ToString()))
            ////                    {
            ////                        // process the special PAs
            ////                        // List all the processess for this PA, then list all the projects for the processess for this PA

            ////                        // Find all projects / support functions that has this practice sampled
            ////                        foreach (var aProcess in OUProcessesList)
            ////                        {
            ////                            // List the process
            ////                            Range processSrcRow = tmpl1Wks.Rows[cTemplateProcessRow];
            ////                            Range processDstRow = aWks.Rows[rowX];
            ////                            processSrcRow.Copy(processDstRow);
            ////                            aWks.Cells[rowX, 2].Value = $"{aPractice.Acronym} {aPractice.Level}.{aPractice.Number}";
            ////                            aWks.Cells[rowX, 3].Value = aProcess.Name;
            ////                            rowX++;


            ////                            // workUnitsInScope contains all the work units, so now add them to the sheet
            ////                            foreach (WorkUnit workUnitToAdd in aProcess.WorkUnits)
            ////                            {
            ////                                // List the work unit in scope
            ////                                Range yelloRow = tmpl1Wks.Rows[cTemplateYellowRow];
            ////                                Range destYellowRow = aWks.Rows[rowX];
            ////                                yelloRow.Copy(destYellowRow);
            ////                                aWks.Cells[rowX, 2].Value = $"{aPractice.Acronym} {aPractice.Level}.{aPractice.Number}";
            ////                                aWks.Cells[rowX, 3].Value = workUnitToAdd.Name;


            ////                                // identify the interviewees
            ////                                List<Schedule2> scheduleForWorkUnit = Schedule2List.Where(x => x.WorkID == workUnitToAdd.ID).ToList();
            ////                                if (scheduleForWorkUnit.Count > 0)
            ////                                {
            ////                                    string meetingParticipantStr = "";
            ////                                    bool firstReview = true;
            ////                                    foreach (var aScheduleItem in scheduleForWorkUnit)
            ////                                    {
            ////                                        if (firstReview)
            ////                                        {
            ////                                            meetingParticipantStr = $"{aScheduleItem.InterviewSession}({aScheduleItem.ParticipantName})";
            ////                                        }
            ////                                        else
            ////                                        {
            ////                                            meetingParticipantStr = meetingParticipantStr + $" {aScheduleItem.InterviewSession}({aScheduleItem.ParticipantName})";
            ////                                        }
            ////                                    }
            ////                                    aWks.Cells[rowX, 8].Value = meetingParticipantStr;
            ////                                }

            ////                                rowX++;

            ////                                Range oeRow = tmpl1Wks.Rows[cOERow];
            ////                                for (int y = 0; y < 1; y++)
            ////                                {
            ////                                    Range destOERow = aWks.Rows[rowX];
            ////                                    oeRow.Copy(destOERow);
            ////                                    aWks.Cells[rowX, 2].Value = workUnitToAdd.Name;
            ////                                    rowX++;
            ////                                }
            ////                            }
            ////                        }
            ////                    }
            ////                }

            ////            }

            ////        }


            ////    }

            ////}
        }

        private const int cDemixOEToolSearchUntilEmptyColumn = 1;
        private const int cDemixOEToolHeadingStartRow = 8;
        private const int cDemixOEToolMaxRows = 1000;


        private void btnDemixTstLinksAndEngl_Click(object sender, EventArgs e)
        {
            // Remove from release 3.8.0.44

            ////// *** Setup the main sheet
            ////// excelApp.Visible = true;

            ////// *** Load main
            //////mainWorkbook = excelApp.Workbooks.Open(persistentData.OEdatabasePathFile);

            ////if ((mainWorkbook = Helper.CheckIfOpenAndOpen(persistentData.DemixToolPathFile)) == null)
            ////{
            ////    //MessageBox.Show($"File {Path.GetFileName(persistentData.OEdatabasePathFile)}" +
            ////    //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.OEdatabasePathFile)}" +
            ////    //    "\n\rdoes not exists");
            ////    MessageBox.Show("File not found, has it been moved or deleted?");
            ////    return;
            ////}
            ////string basePath = Path.GetDirectoryName(persistentData.DemixToolPathFile);

            ////lblStatus.Text = "OEdb:";
            ////string statusStr = "";
            ////foreach (Worksheet wksOEdb in mainWorkbook.Worksheets)
            ////{
            ////    switch (wksOEdb.Name)
            ////    {

            ////        case "CAR":
            ////        case "CM":
            ////        case "DAR":
            ////        case "EST":
            ////        case "MC":
            ////        case "MPM":
            ////        case "OT":
            ////        case "PAD":
            ////        case "PCM":
            ////        case "PLAN":
            ////        case "PQA":
            ////        case "PR":
            ////        case "RDM":
            ////        case "RSK":
            ////        case "VV":
            ////        case "PI":
            ////        case "TS":
            ////        case "GOV":
            ////        case "II":

            ////            //if (wksOEdb.Name=="PI")
            ////            //{
            ////            //    int stop = 1;
            ////            //}
            ////            // *** Find the number of rows
            ////            int NumberOfRows = Helper.FindEndOfWorksheet(wksOEdb, cDemixOEToolSearchUntilEmptyColumn, cDemixOEToolHeadingStartRow, cDemixOEToolMaxRows);
            ////            // Range columnToClear = wksOEdb.Range["Y:Z"];
            ////            // columnToClear.Clear();

            ////            // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
            ////            Range mainRange = wksOEdb.Range["A" + cDemixOEToolHeadingStartRow, "Z" + NumberOfRows];

            ////            // *** List all the hyperlinks https://www.e-iceblue.com/Tutorials/Spire.XLS/Spire.XLS-Program-Guide/Link/Retrieve-Hyperlinks-from-an-Excel-Sheet-in-C-VB.NET.html
            ////            Hyperlinks hyperLinkList = mainRange.Hyperlinks;
            ////            List<Hyperlink> hyperLinksToAdd = new List<Hyperlink>();

            ////            int hyperLinkRow;
            ////            int hyperLinkCol;
            ////            string hyperlinkAddress;
            ////            string PathFileToTest;
            ////            string PathEnglish;

            ////            string englishFullPathFile;
            ////            Boolean fileFound;

            ////            foreach (Hyperlink aHyperlink in hyperLinkList)
            ////            {
            ////                // *** Take each hyperlink and test it
            ////                hyperLinkRow = aHyperlink.Range.Row;
            ////                //if (hyperLinkRow == 9 && wksOEdb.Name == "PI")
            ////                //{
            ////                //    int stop = 1;
            ////                //}
            ////                hyperLinkCol = aHyperlink.Range.Column;
            ////                hyperlinkAddress = aHyperlink.Address;

            ////                // *** Test if the file exists
            ////                fileFound = false;
            ////                PathFileToTest = Path.Combine(basePath, hyperlinkAddress);
            ////                if (File.Exists(PathFileToTest))
            ////                {
            ////                    mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "f"].Value = "ok file";
            ////                    fileFound = true;
            ////                }
            ////                else
            ////                {
            ////                    if (Directory.Exists(PathFileToTest))
            ////                    {
            ////                        mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "f"].Value = "ok directory";
            ////                    }
            ////                    else
            ////                    {
            ////                        mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "f"].Value = "Not ok";
            ////                    }
            ////                }

            ////                // *** Test if the english version exists
            ////                PathEnglish = Path.Combine(Path.GetDirectoryName(PathFileToTest), Path.GetFileNameWithoutExtension(PathFileToTest));
            ////                englishFullPathFile = "";
            ////                string theExtension = Path.GetExtension(PathFileToTest);
            ////                switch (theExtension.ToLower().Trim())
            ////                {
            ////                    case ".xls":
            ////                    case ".xlsx":
            ////                    case ".xlsm":
            ////                        englishFullPathFile = PathEnglish + "-engl.xlsx";
            ////                        break;
            ////                    case ".doc":
            ////                    case ".docx":
            ////                    case ".docm":
            ////                        englishFullPathFile = PathEnglish + "-engl.docx";
            ////                        break;
            ////                    case ".ppt":
            ////                    case ".pptx":
            ////                    case ".pptm":
            ////                        englishFullPathFile = PathEnglish + "-engl.pptx";
            ////                        break;

            ////                }
            ////                if (englishFullPathFile != "")
            ////                {
            ////                    // *** list the new hyperlink
            ////                    if (fileFound && File.Exists(englishFullPathFile))
            ////                    { // file exists, add it
            ////                      // mainRange[hyperLinkRow - cOEDatabaseHeadingStartRow + 1, "z"].Value = "engl";
            ////                        string remainderPath = englishFullPathFile.Remove(0, basePath.Length + 1);
            ////                        string formulaStr = "=hyperlink(\"" + remainderPath + "\",\"engl\")";
            ////                        mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "e"].Formula = formulaStr;

            ////                    }
            ////                    else
            ////                    {
            ////                        mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "e"].Value = "none";
            ////                    }
            ////                }
            ////            }

            ////            // *** Show the status
            ////            statusStr = statusStr + wksOEdb.Name + ".";
            ////            lblStatus.Text = statusStr;
            ////            break;
            ////    }
            ////}
            ////statusStr = statusStr + "done";
            ////lblStatus.Text = statusStr;

            ////MessageBox.Show("Done");

        }

        private void btnImportOE_Click(object sender, EventArgs e)
        {
            // Remove from release 3.8.0.44

            ////OpenFileDialog OEsource = new OpenFileDialog();
            ////OEsource.InitialDirectory = Path.GetDirectoryName(persistentData.DemixTool_ToImport_PathFile);
            ////OEsource.RestoreDirectory = true;
            ////OEsource.Title = "Select source OE file to import";
            ////OEsource.DefaultExt = "*.xlsm";
            ////if (OEsource.ShowDialog() == DialogResult.OK)
            ////{
            ////    lblOEdbSource.Text = OEsource.FileName;
            ////    persistentData.DemixTool_ToImport_PathFile = OEsource.FileName;
            ////    persistentData.SavePersistentData(persistentData);

            ////    lblDemixTool2Import.Text = persistentData.DemixTool_ToImport_PathFile;
            ////}
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void btnDemixOEMerge_Click(object sender, EventArgs e)
        {
            // Remove from release 3.8.0.44

            ////// *** Load source
            //////sourceWorkbook = excelApp.Workbooks.Open(persistentData.AppToolSourcePathFile);
            ////if ((sourceWorkbook = Helper.CheckIfOpenAndOpen(persistentData.DemixTool_ToImport_PathFile)) == null)
            ////{
            ////    //MessageBox.Show($"File {Path.GetFileName(persistentData.AppToolSourcePathFile)}" +
            ////    //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.AppToolSourcePathFile)}" +
            ////    //    "\n\rdoes not exists");
            ////    MessageBox.Show("File not found, has it been moved or deleted?");
            ////    return;
            ////}
            ////// *** Load main
            //////mainWorkbook = excelApp.Workbooks.Open(persistentData.AppToolMainPathFile);
            ////if ((mainWorkbook = Helper.CheckIfOpenAndOpen(persistentData.DemixToolPathFile)) == null)
            ////{
            ////    //MessageBox.Show($"File {Path.GetFileName(persistentData.AppToolMainPathFile)}" +
            ////    //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.AppToolMainPathFile)}" +
            ////    //    "\n\rdoes not exists");
            ////    MessageBox.Show("File not found, has it been moved or deleted?");
            ////    return;
            ////}

            //////  int row;
            //////  string sValueN;
            ////Worksheet wsMain;

            ////lblStatus.Text = "";
            ////string statusStr = "";
            ////foreach (Worksheet wsSource in sourceWorkbook.Sheets)
            ////{

            ////    switch (wsSource.Name)
            ////    {

            ////        case "CAR":
            ////        case "CM":
            ////        case "DAR":
            ////        case "EST":
            ////        case "MC":
            ////        case "MPM":
            ////        case "OT":
            ////        case "PAD":
            ////        case "PCM":
            ////        case "PLAN":
            ////        case "PQA":
            ////        case "PR":
            ////        case "RDM":
            ////        case "RSK":
            ////        case "VV":
            ////        case "PI":
            ////        case "TS":
            ////        case "II":
            ////        case "GOV":

            ////            wsMain = mainWorkbook.Worksheets[wsSource.Name];
            ////            // *** Search string

            ////            //const int cDemixOEToolEmptyColumn = 1;
            ////            //const int cDemixOEToolHeadingStartRow = 8;
            ////            //const int cDemixOEToolMaxRows = 1000;


            ////            // *** Find the number of rows
            ////            int NumberOfRows = Helper.FindEndOfWorksheet(wsSource, cDemixOEToolSearchUntilEmptyColumn, cDemixOEToolHeadingStartRow, cDemixOEToolMaxRows);

            ////            // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
            ////            //ProcessRowsUsingObject(wsMain, wsSource, NumberOfRows, ref statusStr);
            ////            ProcessRowsUsingExcel(wsMain, wsSource, NumberOfRows, ref statusStr);

            ////            break;

            ////    }

            ////}
            ////statusStr = statusStr + "done";
            ////lblStatus.Text = statusStr;

            ////MessageBox.Show("Done");

            ////// Iterate through source, if you find [upload], change to [done-upload] AND copy full line to main

        }

        #region Process Worksheet options
        // *** Option 1 implementation using Ojbect[,] Issue seems to occure if the sheet is filtered
        private void ProcessRowsUsingExcel(Worksheet wsMain, Worksheet wsSource, int NumberOfRows, ref string statusStr)
        {
            //Range soureRange = wsSource.Range["A" + cDemixOEToolHeadingStartRow, "Q" + NumberOfRows];
            //object[,] sourceValue = soureRange.Value;

            // Range mainRange = wsMain.Range["A" + cDemixOEToolHeadingStartRow, "Q" + NumberOfRows];
            //object[,] mainValue = mainRange.Value;

            // *** show progress
            statusStr = statusStr + wsSource.Name + ".";
            lblStatus.Text = statusStr;

            // *** search rows for for upload

            for (int rowS = cDemixOEToolHeadingStartRow; rowS < (NumberOfRows + cDemixOEToolHeadingStartRow); rowS++)
            {
                var cellX = wsSource.Cells[rowS, 17].Value;
                if (cellX != null)
                {
                    string cellXStr = cellX.ToString().Trim().ToUpper();

                    if (cellXStr == "Y" || cellXStr == "YES") // Colum Q has a "Y" or "YES"
                    {
                        // wsMain.Range[wsMain.Cells[rowS, 1], wsMain.Cells[rowS, 16]] = wsSource.Range[wsSource.Cells[rowS, 1], wsSource.Cells[rowS, 16]]; // Rows[rowS];
                        copyRow(wsMain, wsSource, rowS, 1, 16);
                        wsMain.Cells[rowS, 17].Value = DateTime.Now.ToString("s"); // put the short date time here
                        wsSource.Cells[rowS, 17].Value = "updated";
                    }
                }
            }

        }

        private void ProcessRowsUsingObject(Worksheet wsMain, Worksheet wsSource, int NumberOfRows, ref string statusStr)
        {
            Range soureRange = wsSource.Range["A" + cDemixOEToolHeadingStartRow, "Q" + NumberOfRows];
            object[,] sourceValue = soureRange.Value;

            Range mainRange = wsMain.Range["A" + cDemixOEToolHeadingStartRow, "Q" + NumberOfRows];
            object[,] mainValue = mainRange.Value;

            // *** show progress
            statusStr = statusStr + wsSource.Name + ".";
            lblStatus.Text = statusStr;

            // *** search rows for for upload

            for (int rowS = 1; rowS <= (NumberOfRows - cDemixOEToolHeadingStartRow + 1); rowS++)
            {
                if (sourceValue[rowS, 17] != null)
                {
                    string orriginal = sourceValue[rowS, 17].ToString(); // 17 is Q
                    string orriginalUpper = orriginal.ToUpper();

                    if (orriginalUpper == "Y") // Colum Q has a Y
                    {

                        mainValue[rowS, 17] = DateTime.Now.ToString("s"); // s for short date time 
                        for (int col = 9; col <= 16; col++)
                        {
                            mainValue[rowS, col] = sourceValue[rowS, col];
                        }
                        //                                    mainValue[rowS, 9] = "[*Uploaded*]";
                    }
                }
            }

            mainRange.Value = mainValue;
            soureRange.Value = sourceValue;
        }


        private void copyRow(Worksheet wsMain, Worksheet wsSource, int row, int startCol, int endCol)
        {
            for (int aCol = startCol; aCol <= endCol; aCol++)
            {

                if (wsSource.Cells[row, aCol] != null) // && aCol != 4) // do not copy col D
                {
                    // https://docs.devexpress.com/OfficeFileAPI/12235/spreadsheet-document-api/examples/cells/how-to-copy-cell-data-only-cell-style-only-or-cell-data-with-style

                    Range sourceCell = wsSource.Cells[row, aCol];
                    Range destCell = wsMain.Cells[row, aCol];

                    destCell.Value = sourceCell.Value; // .CopyFromRecordset(sourceCell);
                    destCell.Interior.Color = sourceCell.Interior.Color;
                    //destCell.Font.Background = sourceCell.Font.Background;
                    // destCell.Style = sourceCell.Style; // .Value = sourceCell.Value; // .CopyFromRecordset(sourceCell);


                    //destCell.Copy(sourceCell); // .CopyFrom(sourceCell);
                    // wsMain.Cells[row, aCol].Copy(sourceCell);

                    //wsMain.Cells[row, aCol] = wsSource.Cells[row, aCol];

                    //Cell sourceCell = worksheet.Cells["B1"];
                    //sourceCell.Formula = "= PI()";
                    //sourceCell.NumberFormat = "0.0000";
                    //sourceCell.Style = style;
                    //sourceCell.Font.Color = Color.Blue;
                    //sourceCell.Font.Bold = true;
                    //sourceCell.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin);

                    //// Copy all information from the source cell to the "B3" cell. 
                    //worksheet.Cells["A3"].Value = "Copy All";
                    //worksheet.Cells["B3"].CopyFrom(sourceCell);

                }
            }
        }
        #endregion


        private Dictionary<string, string> TmpDicValue = new Dictionary<string, string>();
        private Dictionary<string, string> TmpDictRowCol = new Dictionary<string, string>();

        const int CtmpStartRow = 4; // exclude heading at 3
        const int CtmpEndRow = 35;
        const int CtmpStartCol = 3; // exclude Practice nubmer at 2
        const int CtmpEndCol = 21;


        private void buildTempDictionary()
        {
            TmpDicValue.Clear();
            TmpDictRowCol.Clear();

            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(persistentData.DemixToolPathFile)) == null)
            {
                MessageBox.Show("File not found, has it been moved or deleted?");
                return;
            }
            string basePath = Path.GetDirectoryName(persistentData.DemixToolPathFile);

            lblStatus.Text = "OEdb:";
            //string statusStr = "";

            Worksheet tmpWks = mainWorkbook.Worksheets["tmp"];
            for (int rowX = CtmpStartRow; rowX <= CtmpEndRow; rowX++)
            {
                for (int colY = CtmpStartCol; colY <= CtmpEndCol; colY++)
                {
                    string PAstr = tmpWks.Cells[CtmpStartRow - 1, colY]?.Value?.ToString() ?? "";
                    string numberStr = tmpWks.Cells[rowX, CtmpStartCol - 1]?.Value?.ToString() ?? "";
                    numberStr = numberStr.Replace(',', '.');
                    string KeyStr = PAstr + " " + numberStr;
                    string cellStr = tmpWks.Cells[rowX, colY]?.Value?.ToString() ?? "";
                    TmpDicValue.Add(KeyStr, cellStr);
                    string RowColStr = Helper.GetExcelColumnName(colY) + rowX.ToString();
                    TmpDictRowCol.Add(KeyStr, RowColStr);
                }
            }
        }
        private void btnBuildTmpDictionary_Click(object sender, EventArgs e)
        {



        }

        private void btnBuildOUMaps_Click(object sender, EventArgs e)
        {
            // Remove from release 3.8.0.44
            ////// *** Build temperary dictionary
            ////buildTempDictionary();

            ////// *** Identify pand s files

            ////if ((mainWorkbook = Helper.CheckIfOpenAndOpen(persistentData.DemixToolPathFile)) == null)
            ////{
            ////    MessageBox.Show("File not found, has it been moved or deleted?");
            ////    return;
            ////}
            ////string basePath = Path.GetDirectoryName(persistentData.DemixToolPathFile);

            //////  lblStatus.Text = "OEdb:";

            ////var wksNameArray = WorkUnitList.Where(x => x.ID.Substring(0, 1).ToUpper() == "P" || x.ID.Substring(0, 1).ToUpper() == "S").ToArray();
            //////{ "p1", "p2", "p3", "p4", "p5", "p6", "s1", "s2", "s3", "s4" };
            ////string statusStr = "";

            ////// *** For each p&s build the maps

            ////foreach (var aWksNameX in wksNameArray)
            ////{
            ////    // copy tmp and rename 
            ////    Worksheet projectWks = mainWorkbook.Worksheets["tmp"];
            ////    projectWks.Copy(After: projectWks);
            ////    projectWks = mainWorkbook.Worksheets["tmp (2)"];
            ////    projectWks.Name = aWksNameX.ID;
            ////    projectWks.Cells[1, 1].Value = aWksNameX.Name;

            ////    // setup the links to the detail data
            ////    lblStatus.Text = aWksNameX.ID + "(" + aWksNameX.Name + ")" + "OEdb:";
            ////    statusStr = lblStatus.Text;
            ////    //Worksheet projectWks = mainWorkbook.Worksheets[aWksName];

            ////    foreach (Worksheet wksOEdb in mainWorkbook.Worksheets)
            ////    {
            ////        switch (wksOEdb.Name)
            ////        {

            ////            case "CAR":
            ////            case "CM":
            ////            case "DAR":
            ////            case "EST":
            ////            case "MC":
            ////            case "MPM":
            ////            case "OT":
            ////            case "PAD":
            ////            case "PCM":
            ////            case "PLAN":
            ////            case "PQA":
            ////            case "PR":
            ////            case "RDM":
            ////            case "RSK":
            ////            case "VV":
            ////            case "PI":
            ////            case "TS":
            ////                //  case "GOV":
            ////                //   case "II":

            ////                int NumberOfRows = Helper.FindEndOfWorksheet(wksOEdb, cDemixOEToolSearchUntilEmptyColumn, cDemixOEToolHeadingStartRow, cDemixOEToolMaxRows);
            ////                for (int rowX = cDemixOEToolHeadingStartRow; rowX <= NumberOfRows; rowX++)
            ////                {
            ////                    // Search column B for the key
            ////                    string headingType = wksOEdb.Cells[rowX, 1]?.Value?.ToString().Trim();
            ////                    if (string.Compare(headingType, "4 Prac_Instan", ignoreCase: true) == 0)
            ////                    {
            ////                        // is it the correct project
            ////                        string projectNumber = wksOEdb.Cells[rowX + 1, 2]?.Value?.ToString();
            ////                        if (projectNumber.Substring(0, 2) == projectWks.Name)
            ////                        {
            ////                            string keyStr = wksOEdb.Cells[rowX, 2]?.Value?.ToString().Trim();
            ////                            string rowColStr = FindDictionaryValue(TmpDictRowCol, keyStr);
            ////                            if (!string.IsNullOrEmpty(rowColStr))
            ////                            {
            ////                                //projectWks.Range[rowColStr].Value = wksOEdb.Cells[rowX, 15]?.Value?.ToString() ?? "-";
            ////                                projectWks.Range[rowColStr].Formula = $"={wksOEdb.Name}!O{rowX}"; //=TS!O11
            ////                            }

            ////                        }
            ////                    }

            ////                }



            ////                // *** Show the status
            ////                statusStr = statusStr + wksOEdb.Name + ".";
            ////                lblStatus.Text = statusStr;
            ////                break;
            ////        }
            ////    }
            ////}

            ////statusStr = statusStr + "done";
            ////lblStatus.Text = statusStr;

            ////MessageBox.Show("Done");
        }

        private string FindDictionaryValue(Dictionary<string, string> aDict, string KeyStr)
        {
            // search for the KeyStr
            foreach (var pairVal in aDict)
            {
                if (pairVal.Key == KeyStr)
                    return pairVal.Value;
            }
            return null;

        }


        const int cDXXSearchNumberOfWksRowsCol = 2;
        const int cDMostPAStartRow = 9;
        const int cDMostPAEndRow = 1000;
        const int cDPAtestColumn = 2;
        const int cDMostPAtestOoS = 2;
        const int cDXXWeaknessCol = 12;
        const int cDXXStrengthCol = 13;
        const int cDXXQuestionCol = 11;
        const int cDXXImprovementCol = 14;


        private void btnExtractOEFindings_Click(object sender, EventArgs e)
        {
            // Remove from release 3.8.0.44

            //// *** Load main CMMI tool
            //if ((mainWorkbook = Helper.CheckIfOpenAndOpen(persistentData.DemixToolPathFile)) == null)
            //{
            //    MessageBox.Show("File not found, has it been moved or deleted?");
            //    return;
            //}

            //// *** Does the main workbook contain a findings sheet, if not add one, if it does, assign it and clear it
            //Worksheet findingsWks = AssignOrCreateWorksheet(mainWorkbook, "Findings");
            //findingsWks.Range["A:C"].Clear();
            //findingsWks.Cells[1, 1].Value = "PA";
            //findingsWks.Cells[1, 2].Value = "Strength/Weakness/Improvement";
            //findingsWks.Cells[1, 3].Value = "Description";


            //int findigsRow = 2;

            //lblStatus.Text = "";
            //string statusStr = "";
            //foreach (Worksheet wksMain in mainWorkbook.Worksheets)
            //{
            //    switch (wksMain.Name)
            //    {

            //        case "CAR":
            //        case "CM":
            //        case "DAR":
            //        case "EST":
            //        case "MC":
            //        case "MPM":
            //        case "OT":
            //        case "PAD":
            //        case "PCM":
            //        case "PLAN":
            //        case "PQA":
            //        case "PR":
            //        case "RDM":
            //        case "RSK":
            //        case "VV":
            //        case "PI":
            //        case "TS":
            //        case "GOV":
            //        case "II":
            //            HelperExtractFindingsDemixOE(wksMain, findingsWks, cDXXSearchNumberOfWksRowsCol, cDMostPAStartRow, cDMostPAEndRow, ref findigsRow);
            //            statusStr = statusStr + "." + wksMain.Name;
            //            break;

            //    }
            //    lblStatus.Text = statusStr;
            //}
            ////  wksMain.Application.Visible = true;
            //findingsWks.Activate();
            //MessageBox.Show("Findings extracted");
        }

        private void btnAbridge_Click(object sender, EventArgs e)
        {

            // Remove from release 3.8.0.44

            ////// *** Setup the main sheet
            ////// excelApp.Visible = true;

            ////// *** Load main
            //////mainWorkbook = excelApp.Workbooks.Open(persistentData.OEdatabasePathFile);

            ////if ((mainWorkbook = Helper.CheckIfOpenAndOpen(persistentData.DemixToolPathFile)) == null)
            ////{
            ////    //MessageBox.Show($"File {Path.GetFileName(persistentData.OEdatabasePathFile)}" +
            ////    //    $"\n\rDirectory {Path.GetDirectoryName(persistentData.OEdatabasePathFile)}" +
            ////    //    "\n\rdoes not exists");
            ////    MessageBox.Show("File not found, has it been moved or deleted?");
            ////    return;
            ////}
            ////string basePath = Path.GetDirectoryName(persistentData.DemixToolPathFile);

            ////lblStatus.Text = "OEdb:";
            ////string statusStr = "";
            ////foreach (Worksheet wksOEdb in mainWorkbook.Worksheets)
            ////{
            ////    int fileNumber = 1;

            ////    switch (wksOEdb.Name)
            ////    {

            ////        case "CAR":
            ////        case "CM":
            ////        case "DAR":
            ////        case "EST":
            ////        case "MC":
            ////        case "MPM":
            ////        case "OT":
            ////        case "PAD":
            ////        case "PCM":
            ////        case "PLAN":
            ////        case "PQA":
            ////        case "PR":
            ////        case "RDM":
            ////        case "RSK":
            ////        case "VV":
            ////        case "PI":
            ////        case "TS":
            ////        case "GOV":
            ////        case "II":

            ////            //if (wksOEdb.Name=="PI")
            ////            //{
            ////            //    int stop = 1;
            ////            //}
            ////            // *** Find the number of rows
            ////            int NumberOfRows = Helper.FindEndOfWorksheet(wksOEdb, cDemixOEToolSearchUntilEmptyColumn, cDemixOEToolHeadingStartRow, cDemixOEToolMaxRows);
            ////            // Range columnToClear = wksOEdb.Range["Y:Z"];
            ////            // columnToClear.Clear();

            ////            // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
            ////            Range mainRange = wksOEdb.Range["A" + cDemixOEToolHeadingStartRow, "Z" + NumberOfRows];

            ////            // *** List all the hyperlinks https://www.e-iceblue.com/Tutorials/Spire.XLS/Spire.XLS-Program-Guide/Link/Retrieve-Hyperlinks-from-an-Excel-Sheet-in-C-VB.NET.html
            ////            Hyperlinks hyperLinkList = mainRange.Hyperlinks;
            ////            List<Hyperlink> hyperLinksToAdd = new List<Hyperlink>();

            ////            int hyperLinkRow;
            ////            int hyperLinkCol;
            ////            string hyperlinkAddress;

            ////            foreach (Hyperlink aHyperlink in hyperLinkList)
            ////            {
            ////                // *** Take each hyperlink and test it
            ////                hyperLinkRow = aHyperlink.Range.Row;
            ////                hyperLinkCol = aHyperlink.Range.Column;
            ////                hyperlinkAddress = aHyperlink.Address;

            ////                // *** Test if the file exists

            ////                mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "e"].Value = "engl";
            ////                mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, hyperLinkCol] = mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "e"];
            ////                mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, hyperLinkCol].Value = wksOEdb.Name + fileNumber.ToString("D2");
            ////                fileNumber++;


            ////            }
            ////            foreach (Hyperlink aHyperlink in hyperLinkList)
            ////            {
            ////                hyperLinkRow = aHyperlink.Range.Row;
            ////                hyperLinkCol = aHyperlink.Range.Column;
            ////                //mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, hyperLinkCol].HorizontalAlignment = 
            ////                //mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, hyperLinkCol].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            ////                aHyperlink.Delete();

            ////                // https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.tools.excel.namedrange.font?view=vsto-2017

            ////                wksOEdb.Cells[hyperLinkRow, hyperLinkCol].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            ////                wksOEdb.Cells[hyperLinkRow, hyperLinkCol].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            ////                wksOEdb.Cells[hyperLinkRow, hyperLinkCol].Font.Color = Color.Blue; // https://docs.devexpress.com/OfficeFileAPI/12357/spreadsheet-document-api/examples/formatting/how-to-change-cell-font-and-background-color

            ////                wksOEdb.Cells[hyperLinkRow, hyperLinkCol].Font.UnderLine = true; // https://www.gemboxsoftware.com/spreadsheet/examples/c-sharp-vb-net-excel-style-formatting/202

            ////                // Range aRange = wksOEdb.Range[hyperLinkRow, hyperLinkCol];

            ////            }
            ////            // *** Show the status
            ////            statusStr = statusStr + wksOEdb.Name + ".";
            ////            lblStatus.Text = statusStr;
            ////            break;
            ////    }
            ////}
            ////statusStr = statusStr + "done";
            ////lblStatus.Text = statusStr;

            ////MessageBox.Show("Done");
        }

        private void tabDemixTool_Click(object sender, EventArgs e)
        {

        }

        private void lbStatCASPlanLoaded_Click(object sender, EventArgs e)
        {

        }

        private void btnBuildPandS_Click(object sender, EventArgs e)
        {

            int i = 1;
        }

        private void btnOpenBaseCASPlan_Click(object sender, EventArgs e)
        {
            //CASFileObject.LoadPersistant();
            if (CASFileObject.SelectFileToLoad(TargetFileObject.CCASinName) == false)
            {
                MessageBox.Show($"No file selected.");
            }
            else
            {
                CASFileObject.SavePersistant(CASFileObject);
            }
        }

        private void btnReloadCASPlan_Click(object sender, EventArgs e)
        {
            //aWorkbook = excelApp.Workbooks.Open(LblSourceFilePlan2.Text.ToString());
            if (CASFileObject.LoadCASFile() == true)
            { // file was loaded
                CASFileObject.SavePersistant(CASFileObject);
            }
            else
            { // file was not loaded
            }


            return;


        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (CASFileObject.CreateSchedule1() == true)
            { // Generated the schedule

            }
            else
            {// Could not generate schedule

            }
        }

        private void btnReloadSchedule2AndGenerateCASSheets_Click(object sender, EventArgs e)
        {
            //CASFileObject.ReloadSchedule2();
            //bool insertRole = ;
            if (CASFileObject.Generate_OUParticipants(chkInsertRole.Checked)) // This includes reloading it
            { // All ok

            }
            else
            {
                MessageBox.Show("Error reloading schedule 2 and generating CAS sheets!");
            }
        }

        private void btnGenerating_SupportAndProjectCASSheets(object sender, EventArgs e)
        {
            if (CASFileObject.Generate_SupportAndProjectCASSheets())
            {
                // All ok
            }
            else
            {
                MessageBox.Show("Error generating Support and Project CAS sheets!");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //CASFileObject.LoadPersistant();
            if (CASOEdbObject.SelectFileToLoad(TargetFileObject.COEdbinName) == false)
            {
                MessageBox.Show($"No file selected.");
            }
            else
            {
                CASOEdbObject.SavePersistant(CASOEdbObject);
            }
        }

        private void btnSelectQuestionAndModel2_Click(object sender, EventArgs e)
        {
            //BASEQuestionObject 
            if (BASEQuestionObject.SelectFileToLoad(TargetFileObject.CQuestionInName) == false)
            {
                MessageBox.Show($"No file selected.");
            }
            else
            {


                BASEQuestionObject.SavePersistant(BASEQuestionObject);
            }
        }

        private void btnReloadQuestionsAndModel2_Click(object sender, EventArgs e)
        {
            if (BASEQuestionObject.LoadTheQuestionAndModelFile(lblStatus) == false)
            {
                MessageBox.Show($"Could not read the questions and model file.");
            }
            else
            {

                BASEQuestionObject.SavePersistant(BASEQuestionObject);
            }
        }

        private void btnGenerateOEdb2_Click(object sender, EventArgs e)
        {
            //    private TargetCASFileObject CASFileObject;
            //private TargetOEFileObject CASOEdbObject;
            //private TargetQuestionsFileObject BASEQuestionObject;

            if (CASOEdbObject.GenerateFullOEdb2(CASFileObject, BASEQuestionObject) == false)
            {
                MessageBox.Show($"Could not complete the OE database generation!");
            }
            ;
        }

        private void btnTestLinksAndEngl2_Click(object sender, EventArgs e)
        {
            if (CASOEdbObject.TestLinksAndEnglish2(lblStatus) == false)

            {
                MessageBox.Show($"Could not complete the link testing task!");
            }
        }

        private void btnExtractFindings2_Click(object sender, EventArgs e)
        {
            if (CASOEdbObject.ExtractOEFindings2(lblStatus) == false)

            {
                MessageBox.Show($"Could not complete finding extraction!");
            }
        }

        private void btnBuildOUMaps2_Click(object sender, EventArgs e)
        {
            if (CASOEdbObject.BuildOUMaps2(lblStatus, CASFileObject) == false)

            {
                MessageBox.Show($"Could not build the OU maps!");
            }
        }

        private void btnBuildAbridged2_Click(object sender, EventArgs e)
        {

            if (CASOEdbObject.BuildAbridgedOEdb2(lblStatus) == false)

            {
                MessageBox.Show($"Could not build the abridged OEdb");
            }
        }

        private void btnImportOEdb2_Click(object sender, EventArgs e)
        {
            //CASFileObject.LoadPersistant();
            if (CASOEdbImportObject.SelectFileToLoad(TargetFileObject.COEdbATMinName) == false)
            {
                MessageBox.Show($"No file selected.");
            }
            else
            {
                CASOEdbImportObject.SavePersistant(CASOEdbImportObject);
            }
        }

        private void btnMergeATMintoATL2_Click(object sender, EventArgs e)
        {
            if (CASOEdbImportObject == null)
            {
                MessageBox.Show($"The import OEdbATM has not been selected");
                return;
            }
            if (CASOEdbObject == null)
            {
                MessageBox.Show($"The main OEdbATL file has not been selected");
                return;
            }

            if (CASOEdbObject.MergeATMintoATL2(lblStatus, CASOEdbImportObject) == false)
            {
                MessageBox.Show($"The mergining has not been completed.");
            }

        }

        private void btnSelectXlsxAdmin_Click(object sender, EventArgs e)
        {
            
            if (BASEDataReferenceObject.SelectFileToLoad("Data_Reference") == false)
            {
                MessageBox.Show($"No file selected.");
            }
            else
            {
                BASEDataReferenceObject.SavePersistant(BASEDataReferenceObject);
                //BASEPresentationObject.ClearPathFile();
            }
        }


        private void btnSelectPptxAdmin_Click(object sender, EventArgs e)
        {
            if (BASEPresentationObject.SelectFileToLoad("") == false)
            {
                MessageBox.Show($"No file selected.");
            }
            else
            {
                BASEPresentationObject.SavePersistant(BASEPresentationObject);
                //BASEDataReferenceObject.ClearPathFile();
            }
        }
        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            // *** Open and update
            //aWorkbook = excelApp.Workbooks.Open(LblSourceFilePlan2.Text.ToString());
            string presentationPath = Path.GetDirectoryName(BASEPresentationObject._directoryFileName);
            string dataReferencePath = Path.GetDirectoryName(BASEDataReferenceObject._directoryFileName);
                
            if (presentationPath != dataReferencePath)
            {
                var resultX = MessageBox.Show($"The data reference and presentation files are in different directories\n" +
                    $"Data={dataReferencePath}\nPresentation={presentationPath}\nPress [Yes] to continue and [No] to stop and first correct.", caption: "Warning", MessageBoxButtons.YesNo);
                if (resultX != DialogResult.Yes) return;
            }
            if (BASEPresentationObject.UpdateLinks(BASEDataReferenceObject) == true)
            { // file was loaded
                BASEPresentationObject.SavePersistant(BASEPresentationObject);
                MessageBox.Show($"Links updated!");
            }
            else
            { // file was not loaded
                MessageBox.Show($"Link updates unsuccessfull!");
            }

        }

        private void WorkinWritingToZip()
        {
            // *** Open and update

            string fileNameNoExt = Path.GetFileNameWithoutExtension(BASEPresentationObject._directoryFileName);
            string fileExt = Path.GetExtension(BASEPresentationObject._directoryFileName);
            string directoryName = Path.GetDirectoryName(BASEPresentationObject._directoryFileName);

            string zipFileName = Path.Combine(directoryName, fileNameNoExt + ".zip");

            File.Copy(BASEPresentationObject._directoryFileName, zipFileName);
            

                //Path.Combine(Path.GetDirectoryName(BASEDataReferenceObject._directoryFileName),
                //"myZip.zip");

            // https://docs.telerik.com/devtools/document-processing/libraries/radziplibrary/features/update-ziparchive
            try
            {
                using (Stream stream = File.Open(zipFileName, FileMode.Open))
                {
                    using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Update, false, null))
                    {
                        // Display the list of the files in the selected zip file using the ZipArchive.Entries property. 
                        foreach (ZipArchiveEntry zEntry in archive.Entries)
                        {
                            string aStr = zEntry.FullName;
                        }

                        // *** Add entry
                        ZipArchiveEntry aNewEntry1 = archive.CreateEntry("myprogramEntry.txt");
                        ZipArchiveEntry aNewEntry2 = archive.GetEntry("text.txt");
                        if (aNewEntry2 == null) aNewEntry2 = archive.CreateEntry("text.txt");

                        // *** Delete entry
                        ZipArchiveEntry addedEntry1 = archive.GetEntry("myprogramEntry.txt");
                        addedEntry1.Delete();


                        // *** Update entry
                        ZipArchiveEntry entry2 = archive.GetEntry("text.txt");
                        if (entry2 != null)
                        {
                            Stream entryStream = entry2.Open();
                            StreamReader reader = new StreamReader(entryStream);
                            string content = reader.ReadToEnd();
                            string contentReplaced = content.Replace("line", "<replaced line>");
                            if (string.IsNullOrEmpty(contentReplaced)) { contentReplaced = "My line to insert."; }
                            //entryStream.Seek(0, SeekOrigin.End);
                            entryStream.Seek(0, SeekOrigin.Begin);
                            StreamWriter writer = new StreamWriter(entryStream);
                            writer.WriteLine(contentReplaced);
                            writer.Flush();
                        }

                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Message:{ex.Message}");
            }
        }
    }
}
