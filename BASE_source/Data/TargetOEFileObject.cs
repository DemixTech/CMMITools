using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace BASE.Data
{
    [Serializable]
    public class TargetOEFileObject : TargetFileObject
    {

        const int cDemixOEToolSearchUntilEmptyColumn = 1;
        const int cDemixOEToolHeadingStartRow = 8;
        const int cDemixOEToolMaxRows = 1000;

        const int cDXXSearchNumberOfWksRowsCol = 2;
        const int cDMostPAStartRow = 9;
        const int cDMostPAEndRow = 1000;

        const int CD_Heading = 1;
        const int CD_practiceCol = 2;
        const int CD_weaknessCol = 12;
        const int CD_strengthCol = 13;
        const int CD_recommendationCol = 14;

        private Dictionary<string, string> TmpDicValue = new Dictionary<string, string>();
        private Dictionary<string, string> TmpDictRowCol = new Dictionary<string, string>();

        const int CtmpStartRow = 4; // exclude heading at 3
        const int CtmpEndRow = 35;
        const int CtmpStartCol = 3; // exclude Practice nubmer at 2
        const int CtmpEndCol = 21;

        //public override bool LoadFileExcelFileData(string fileNameKeyWord)
        //{
        //    throw new NotImplementedException();
        //}
        public string OEdataStr = "OE data string";

        public override bool LoadPersistantXMLdata()
        {
            try
            {
                // base.LoadPersistant(); override the base function, to load all information from here for this object and its parent
                if (File.Exists(_directoryFileNameXML))
                {
                    // If the directory and file name exists, laod the data
                    var xs = new XmlSerializer(typeof(TargetOEFileObject)); // TargetCASFileObject));
                    using (FileStream xmlLoad = File.Open(_directoryFileNameXML, FileMode.Open))
                    {
                        var pData = (TargetOEFileObject)xs.Deserialize(xmlLoad);
                        this.DirectoryFileName = pData._directoryFileName;

                        // *** Load the object elements belwo
                        this.OEdataStr = pData.OEdataStr;
                    }
                    return true; // loaded successfull
                }
                else
                {
                    return false; //load unsuccessfull
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error {ex.Message} loading {_directoryFileNameXML}");
                return false;
            }

        }

        public override void SavePersistant(object o)
        {
            if (o is TargetOEFileObject tc)
            {
                if (!Directory.Exists(Path.GetDirectoryName(_directoryFileNameXML)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(_directoryFileNameXML)); ;
                }

                var xs = new XmlSerializer(typeof(TargetOEFileObject));
                using (FileStream stream = File.Create(_directoryFileNameXML))
                {
                    xs.Serialize(stream, tc);
                }

            }
            else
            {
                throw new NotImplementedException("Object missmatched");

            }
        }

        public bool GenerateFullOEdb2(TargetCASFileObject CASFileObject2, TargetQuestionsFileObject BASEQuestionObject2)
        {
            DialogResult dialogResult = MessageBox.Show("Make sure Processess are correcly listed in tab:Project&Support! Continue?", "Warning", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                //do something
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
                return false;
            }

            string[] mostPAs = { "PI", "TS", "PQA", "PR", "RDM", "VV", "MPM", "PAD", "PCM", "RSK", "OT", "EST", "MC", "PLAN", "CAR", "CM", "DAR", "SAM" };
            string[] specialPAs = { "GOV", "II" };
            const int cTemplateLevelRow = 3;
            const int cTemplateBlueRow = 4;
            const int cTemplateProcessRow = 5;
            const int cTemplateYellowRow = 6;
            const int cOERow = 7;

            // open demix tool, if not open
            Workbook demixToolWkb;
            if ((demixToolWkb = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
            {
                MessageBox.Show("Cannot open the demix tool, is the file moved or deleted?");
                return false;
            }

            // generic variables
            Worksheet tmpl1Wks = demixToolWkb.Worksheets["Template1"];

            // demixToolWkb contains the opened workbook
            foreach (PracticeArea aPracticeArea in BASEQuestionObject2.CMMIModel2)
            {
                // DEBUG CODE, SKIP most PAs
                //if (mostPAs.Contains(aPracticeArea.PAcode.ToString())) continue;
                // create a worksheet if it does not exist
                //Worksheet aWks = Helper.OpenOrElseCreateWks(demixToolWkb, aPracticeArea.PAcode.ToString());
                foreach (Worksheet findWks in demixToolWkb.Worksheets)
                {
                    if (findWks.Name == aPracticeArea.PAcode.ToString()) findWks.Delete();
                }
                // Copy the template2 over that worksheet
                Worksheet sourceWks;
                Worksheet aWks;
                sourceWks = demixToolWkb.Worksheets["Template2"];
                //aWks = demixToolWkb.Worksheets.Add();
                int numberOfWks = demixToolWkb.Worksheets.Count;
                sourceWks.Copy(After: demixToolWkb.Worksheets[numberOfWks]);
                aWks = demixToolWkb.Worksheets[numberOfWks + 1];
                aWks.Name = aPracticeArea.PAcode.ToString();

                // Setup the headings
                aWks.Cells[1, 1].Value = aPracticeArea.Name;
                aWks.Cells[2, 1].Value = aPracticeArea.NameChinese;
                aWks.Cells[3, 2].Value = aPracticeArea.Intent;
                aWks.Cells[4, 2].Value = aPracticeArea.IntentChinese;
                aWks.Cells[5, 2].Value = aPracticeArea.Value;
                aWks.Cells[6, 2].Value = aPracticeArea.ValueChinese;

                // Setup parameters
                int rowX = 9; // the starting row to process
                              // Build each of the levels 
                for (int levelX = 1; levelX <= 5; levelX++)
                {

                    // Find all practices at this level
                    var levelPractices =
                        from aPractice in aPracticeArea.Practices
                        where aPractice.Level == levelX
                        orderby aPractice.Number
                        select aPractice;

                    if (levelPractices?.Count() >= 1)
                    {
                        // Practices found for this level
                        // Copy the level
                        Range levelRow = tmpl1Wks.Rows[cTemplateLevelRow];
                        Range destLevelRow = aWks.Rows[rowX];
                        levelRow.Copy(destLevelRow);

                        // Set the level number
                        aWks.Cells[rowX, 2].Value = $"Level {levelX}";
                        rowX++;

                        // run through each practice and process it
                        foreach (Practice aPractice in levelPractices)
                        {
                            // Copy the practice heading
                            Range blueRow = tmpl1Wks.Rows[cTemplateBlueRow];
                            Range destBlueRow = aWks.Rows[rowX];
                            blueRow.Copy(destBlueRow);
                            aWks.Cells[rowX, 2].Value = $"{aPractice.Acronym} {aPractice.Level}.{aPractice.Number}";
                            // Extract statement, work products, activities and questions
                            string statementStr, workProductStr, activityStr, questionStr;
                            Helper.ExtractPracticeAreaInformation(aPractice, out statementStr, out workProductStr,
                                out activityStr, out questionStr);
                            aWks.Cells[rowX, 3].Value = statementStr;
                            aWks.Cells[rowX, 9].Value = workProductStr;
                            aWks.Cells[rowX, 10].Value = activityStr;
                            aWks.Cells[rowX, 11].Value = questionStr;

                            rowX++;

                            if (mostPAs.Contains(aPracticeArea.PAcode.ToString()))
                            {
                                // process most PAs
                                // Find all projects / support funcitons that has this practice sampled
                                List<WorkUnit> workUnitsInScope = new List<WorkUnit>();
                                foreach (WorkUnit aWorkUnit in CASFileObject2.WorkUnitList2)
                                {
                                    // See if the practice is present in the work unit practice list
                                    var matchingPAList = from aPAitem in aWorkUnit.PAlist
                                                         where aPAitem.PAcode == aPracticeArea.PAcode
                                                         select aPAitem;
                                    // If it is present, add it to the list
                                    if (matchingPAList?.Count() > 0)
                                    {
                                        workUnitsInScope.Add(aWorkUnit);
                                    }
                                }

                                // workUnitsInScope contains all the work units, so now add them to the sheet
                                foreach (WorkUnit workUnitToAdd in workUnitsInScope)
                                {
                                    // List the work unit in scope
                                    Range yelloRow = tmpl1Wks.Rows[cTemplateYellowRow];
                                    Range destYellowRow = aWks.Rows[rowX];
                                    yelloRow.Copy(destYellowRow);
                                    aWks.Cells[rowX, 2].Value = $"{aPractice.Acronym} {aPractice.Level}.{aPractice.Number}";
                                    aWks.Cells[rowX, 3].Value = workUnitToAdd.Name;

                                    // identify the interviewees
                                    List<Schedule2> scheduleForWorkUnit = CASFileObject2.Schedule2List2.Where(x => x.PA == aPracticeArea.PAcode.ToString() && x.WorkID == workUnitToAdd.ID).ToList();
                                    if (scheduleForWorkUnit.Count > 0)
                                    {
                                        string meetingParticipantStr = "";
                                        bool firstReview = true;
                                        foreach (var aScheduleItem in scheduleForWorkUnit)
                                        {
                                            if (firstReview)
                                            {
                                                meetingParticipantStr = $"{aScheduleItem.InterviewSession}({aScheduleItem.ParticipantName})";
                                            }
                                            else
                                            {
                                                meetingParticipantStr = meetingParticipantStr + $" {aScheduleItem.InterviewSession}({aScheduleItem.ParticipantName})";
                                            }
                                        }
                                        aWks.Cells[rowX, 8].Value = meetingParticipantStr;
                                    }
                                    // List staff applicable to this project
                                    // var staffForThisWorkUnit = StaffList.Where(x => x.WorkID == workUnitToAdd.ID).ToList();

                                    //var listOfInterviewees = StaffList.Where(x => x.)
                                    rowX++;

                                    Range oeRow = tmpl1Wks.Rows[cOERow];
                                    for (int y = 0; y < 2; y++)
                                    {
                                        Range destOERow = aWks.Rows[rowX];
                                        oeRow.Copy(destOERow);
                                        aWks.Cells[rowX, 2].Value = workUnitToAdd.Name;
                                        rowX++;
                                    }
                                }

                            }
                            else
                            {
                                if (specialPAs.Contains(aPracticeArea.PAcode.ToString()))
                                {
                                    // process the special PAs
                                    // List all the processess for this PA, then list all the projects for the processess for this PA

                                    // Find all projects / support functions that has this practice sampled
                                    foreach (var aProcess in CASFileObject2.OUProcessesList2)
                                    {
                                        // List the process
                                        Range processSrcRow = tmpl1Wks.Rows[cTemplateProcessRow];
                                        Range processDstRow = aWks.Rows[rowX];
                                        processSrcRow.Copy(processDstRow);
                                        aWks.Cells[rowX, 2].Value = $"{aPractice.Acronym} {aPractice.Level}.{aPractice.Number}";
                                        aWks.Cells[rowX, 3].Value = aProcess.Name;
                                        rowX++;


                                        // workUnitsInScope contains all the work units, so now add them to the sheet
                                        foreach (WorkUnit workUnitToAdd in aProcess.WorkUnits)
                                        {
                                            // List the work unit in scope
                                            Range yelloRow = tmpl1Wks.Rows[cTemplateYellowRow];
                                            Range destYellowRow = aWks.Rows[rowX];
                                            yelloRow.Copy(destYellowRow);
                                            aWks.Cells[rowX, 2].Value = $"{aPractice.Acronym} {aPractice.Level}.{aPractice.Number}";
                                            aWks.Cells[rowX, 3].Value = workUnitToAdd.Name;


                                            // identify the interviewees
                                            List<Schedule2> scheduleForWorkUnit = CASFileObject2.Schedule2List2.Where(x => x.WorkID == workUnitToAdd.ID).ToList();
                                            if (scheduleForWorkUnit.Count > 0)
                                            {
                                                string meetingParticipantStr = "";
                                                bool firstReview = true;
                                                foreach (var aScheduleItem in scheduleForWorkUnit)
                                                {
                                                    if (firstReview)
                                                    {
                                                        meetingParticipantStr = $"{aScheduleItem.InterviewSession}({aScheduleItem.ParticipantName})";
                                                    }
                                                    else
                                                    {
                                                        meetingParticipantStr = meetingParticipantStr + $" {aScheduleItem.InterviewSession}({aScheduleItem.ParticipantName})";
                                                    }
                                                }
                                                aWks.Cells[rowX, 8].Value = meetingParticipantStr;
                                            }

                                            rowX++;

                                            Range oeRow = tmpl1Wks.Rows[cOERow];
                                            for (int y = 0; y < 1; y++)
                                            {
                                                Range destOERow = aWks.Rows[rowX];
                                                oeRow.Copy(destOERow);
                                                aWks.Cells[rowX, 2].Value = workUnitToAdd.Name;
                                                rowX++;
                                            }
                                        }
                                    }
                                }
                            }

                        }

                    }


                }

            }

            return true;
        }

        public bool TestLinksAndEnglish2(System.Windows.Forms.Label lblStatus)
        {

            // *** Setup the main sheet
            // excelApp.Visible = true;

            // *** Load main
            //mainWorkbook = excelApp.Workbooks.Open(persistentData.OEdatabasePathFile);
            Workbook mainWorkbook;
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
            {
                MessageBox.Show("File not found, has it been moved or deleted?");
                return false;
            }
            string basePath = Path.GetDirectoryName(_directoryFileName);

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
                        int NumberOfRows = Helper.FindEndOfWorksheet(wksOEdb, cDemixOEToolSearchUntilEmptyColumn, cDemixOEToolHeadingStartRow, cDemixOEToolMaxRows);
                        // Range columnToClear = wksOEdb.Range["Y:Z"];
                        // columnToClear.Clear();

                        // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                        Range mainRange = wksOEdb.Range["A" + cDemixOEToolHeadingStartRow, "Z" + NumberOfRows];

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
                            //if (hyperLinkRow == 9 && wksOEdb.Name == "PI")
                            //{
                            //    int stop = 1;
                            //}
                            hyperLinkCol = aHyperlink.Range.Column;
                            hyperlinkAddress = aHyperlink.Address;

                            // *** Test if the file exists
                            fileFound = false;
                            PathFileToTest = Path.Combine(basePath, hyperlinkAddress);
                            if (File.Exists(PathFileToTest))
                            {
                                mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "f"].Value = "ok file";
                                fileFound = true;
                            }
                            else
                            {
                                if (Directory.Exists(PathFileToTest))
                                {
                                    mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "f"].Value = "ok directory";
                                }
                                else
                                {
                                    mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "f"].Value = "Not ok";
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
                                    mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "e"].Formula = formulaStr;

                                }
                                else
                                {
                                    mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "e"].Value = "none";
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


            return true;
        }


        public bool ExtractOEFindings2(System.Windows.Forms.Label lblStatus)
        {
            // *** Load main CMMI tool
            Workbook mainWorkbook;
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
            {
                MessageBox.Show("File not found, has it been moved or deleted?");
                return false;
            }

            // *** Does the main workbook contain a findings sheet, if not add one, if it does, assign it and clear it
            Worksheet findingsWks = Helper.AssignOrCreateWorksheet(mainWorkbook, "Findings", "Processes");
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
                    case "GOV":
                    case "II":
                        HelperExtractFindingsDemixOE(wksMain, findingsWks, cDXXSearchNumberOfWksRowsCol, cDMostPAStartRow, cDMostPAEndRow, ref findigsRow);
                        statusStr = statusStr + "." + wksMain.Name;
                        break;



                }
                lblStatus.Text = statusStr;
            }
            //  wksMain.Application.Visible = true;
            findingsWks.Activate();
            MessageBox.Show("Findings extracted");
            return true;
        }

        public bool BuildOUMaps2(System.Windows.Forms.Label lblStatus, TargetCASFileObject CASFileObject2)
        {
            // *** Build temperary dictionary
            buildTempDictionary();

            // *** Identify pand s files
            Workbook mainWorkbook;
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
            {
                MessageBox.Show("File not found, has it been moved or deleted?");
                return false;
            }
            //string basePath = Path.GetDirectoryName(_directoryFileName);

            //  lblStatus.Text = "OEdb:";

            var wksNameArray = CASFileObject2.WorkUnitList2.Where(x => x.ID.Substring(0, 1).ToUpper() == "P" || x.ID.Substring(0, 1).ToUpper() == "S").ToArray();
            //{ "p1", "p2", "p3", "p4", "p5", "p6", "s1", "s2", "s3", "s4" };
            string statusStr = "";

            // *** For each p&s build the maps

            foreach (var aWksNameX in wksNameArray)
            {
                // copy tmp and rename 
                Worksheet projectWks = mainWorkbook.Worksheets["tmp"];
                projectWks.Copy(After: projectWks);
                projectWks = mainWorkbook.Worksheets["tmp (2)"];
                projectWks.Name = aWksNameX.ID;
                projectWks.Cells[1, 1].Value = aWksNameX.Name;

                // setup the links to the detail data
                lblStatus.Text = aWksNameX.ID + "(" + aWksNameX.Name + ")" + "OEdb:";
                statusStr = lblStatus.Text;
                //Worksheet projectWks = mainWorkbook.Worksheets[aWksName];

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
                            //  case "GOV":
                            //   case "II":

                            int NumberOfRows = Helper.FindEndOfWorksheet(wksOEdb, cDemixOEToolSearchUntilEmptyColumn, cDemixOEToolHeadingStartRow, cDemixOEToolMaxRows);
                            for (int rowX = cDemixOEToolHeadingStartRow; rowX <= NumberOfRows; rowX++)
                            {
                                // Search column B for the key
                                string headingType = wksOEdb.Cells[rowX, 1]?.Value?.ToString().Trim();
                                if (string.Compare(headingType, "4 Prac_Instan", ignoreCase: true) == 0)
                                {
                                    // is it the correct project
                                    string projectNumber = wksOEdb.Cells[rowX + 1, 2]?.Value?.ToString();
                                    if (projectNumber.Substring(0, 2) == projectWks.Name)
                                    {
                                        string keyStr = wksOEdb.Cells[rowX, 2]?.Value?.ToString().Trim();
                                        string rowColStr = FindDictionaryValue(TmpDictRowCol, keyStr);
                                        if (!string.IsNullOrEmpty(rowColStr))
                                        {
                                            //projectWks.Range[rowColStr].Value = wksOEdb.Cells[rowX, 15]?.Value?.ToString() ?? "-";
                                            projectWks.Range[rowColStr].Formula = $"={wksOEdb.Name}!O{rowX}"; //=TS!O11
                                        }

                                    }
                                }

                            }



                            // *** Show the status
                            statusStr = statusStr + wksOEdb.Name + ".";
                            lblStatus.Text = statusStr;
                            break;
                    }
                }
            }

            statusStr = statusStr + "done";
            lblStatus.Text = statusStr;

            MessageBox.Show("Done");
            return true;

        }

        public bool BuildAbridgedOEdb2(System.Windows.Forms.Label lblStatus)
        {
            // *** Check if the file exists, if it does not, copy it and then abridge it
            string fileNameNoExt = Path.GetFileNameWithoutExtension(_directoryFileName);
            string fileExt = Path.GetExtension(_directoryFileName);
            string pathSTr = Path.GetDirectoryName(_directoryFileName);
            string abridgedFileName;
            int counter = 1;
            do
            {
                abridgedFileName = Path.Combine(pathSTr, fileNameNoExt + counter.ToString() + fileExt);
                counter++;
            } while (File.Exists(abridgedFileName));

            // At this point the abridgedFileName should not exist, so copy it
            File.Copy(_directoryFileName, abridgedFileName);

            // Now process the abridged filename
            Workbook mainWorkbook;
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(abridgedFileName)) == null)
            {
                MessageBox.Show("File not found, has it been moved or deleted?");
                return false;
            }
           // string basePath = Path.GetDirectoryName(persistentData.DemixToolPathFile);

            lblStatus.Text = "OEdb:";
            string statusStr = "";
            foreach (Worksheet wksOEdb in mainWorkbook.Worksheets)
            {
                int fileNumber = 1;

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
                        int NumberOfRows = Helper.FindEndOfWorksheet(wksOEdb, cDemixOEToolSearchUntilEmptyColumn, cDemixOEToolHeadingStartRow, cDemixOEToolMaxRows);
                        // Range columnToClear = wksOEdb.Range["Y:Z"];
                        // columnToClear.Clear();

                        // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                        Range mainRange = wksOEdb.Range["A" + cDemixOEToolHeadingStartRow, "Z" + NumberOfRows];

                        // *** List all the hyperlinks https://www.e-iceblue.com/Tutorials/Spire.XLS/Spire.XLS-Program-Guide/Link/Retrieve-Hyperlinks-from-an-Excel-Sheet-in-C-VB.NET.html
                        Hyperlinks hyperLinkList = mainRange.Hyperlinks;
                        List<Hyperlink> hyperLinksToAdd = new List<Hyperlink>();

                        int hyperLinkRow;
                        int hyperLinkCol;
                        string hyperlinkAddress;

                        foreach (Hyperlink aHyperlink in hyperLinkList)
                        {
                            // *** Take each hyperlink and test it
                            hyperLinkRow = aHyperlink.Range.Row;
                            hyperLinkCol = aHyperlink.Range.Column;
                            hyperlinkAddress = aHyperlink.Address;

                            // *** Test if the file exists

                            mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "e"].Value = "engl";
                            mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, hyperLinkCol] = mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, "e"];
                            mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, hyperLinkCol].Value = wksOEdb.Name + fileNumber.ToString("D2");
                            fileNumber++;


                        }
                        foreach (Hyperlink aHyperlink in hyperLinkList)
                        {
                            hyperLinkRow = aHyperlink.Range.Row;
                            hyperLinkCol = aHyperlink.Range.Column;
                            //mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, hyperLinkCol].HorizontalAlignment = 
                            //mainRange[hyperLinkRow - cDemixOEToolHeadingStartRow + 1, hyperLinkCol].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            aHyperlink.Delete();

                            // https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.tools.excel.namedrange.font?view=vsto-2017

                            wksOEdb.Cells[hyperLinkRow, hyperLinkCol].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            wksOEdb.Cells[hyperLinkRow, hyperLinkCol].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            wksOEdb.Cells[hyperLinkRow, hyperLinkCol].Font.Color = Color.Blue; // https://docs.devexpress.com/OfficeFileAPI/12357/spreadsheet-document-api/examples/formatting/how-to-change-cell-font-and-background-color

                            wksOEdb.Cells[hyperLinkRow, hyperLinkCol].Font.UnderLine = true; // https://www.gemboxsoftware.com/spreadsheet/examples/c-sharp-vb-net-excel-style-formatting/202

                            // Range aRange = wksOEdb.Range[hyperLinkRow, hyperLinkCol];

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
            return true;
        }

        public bool MergeATMintoATL2(System.Windows.Forms.Label lblStatus, TargetOEFileObject fileToImport)
        {

            // *** Load source
            Workbook sourceWorkbook;
            if ((sourceWorkbook = Helper.CheckIfOpenAndOpenXlsx(fileToImport._directoryFileName)) == null)
            {
                MessageBox.Show("File not found, has it been moved or deleted?");
                return false;
            }
            // *** Load main
            Workbook mainWorkbook;
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
            {
                MessageBox.Show("File not found, has it been moved or deleted?");
                return false;
            }

            //  string sValueN;
            Worksheet wsMain;

            lblStatus.Text = "";
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
                    case "II":
                    case "GOV":

                        wsMain = mainWorkbook.Worksheets[wsSource.Name];
                        // *** Find the number of rows
                        int NumberOfRows = Helper.FindEndOfWorksheet(wsSource, cDemixOEToolSearchUntilEmptyColumn, cDemixOEToolHeadingStartRow, cDemixOEToolMaxRows);

                        // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                        //ProcessRowsUsingObject(wsMain, wsSource, NumberOfRows, ref statusStr);
                        ProcessRowsUsingExcel(wsMain, wsSource, NumberOfRows);
                        lblStatus.Text = lblStatus.Text + " " + wsSource.Name;
                        break;

                }

            }
            lblStatus.Text = lblStatus.Text + "done";
            MessageBox.Show("Done");
            return true;

        }


        // ********** HELPER METHODS ****************

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

        private void buildTempDictionary()
        {
            TmpDicValue.Clear();
            TmpDictRowCol.Clear();

            Workbook mainWorkbook;
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
            {
                MessageBox.Show("File not found, has it been moved or deleted?");
                return;
            }
            //string basePath = Path.GetDirectoryName(persistentData.DemixToolPathFile);

            //lblStatus.Text = "OEdb:";
            //string statusStr = "";

            Worksheet tmpWks = Helper.FindWorksheet(mainWorkbook, "tmp");
            if (tmpWks == null) MessageBox.Show($"The tmp worksheet could not be found. Use the latest-new OEdb template!");
            //mainWorkbook.Worksheets["tmp"];
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

        private void ProcessRowsUsingExcel(Worksheet wsMain, Worksheet wsSource, int NumberOfRows)
        {
            // *** Clear the columns
            Range wsMainToClear = wsMain.Range[wsMain.Cells[9, 17], wsMain.Cells[NumberOfRows, 17]];
            wsMainToClear.Interior.Color = Color.White;
            
            Range wsImportToClear = wsSource.Range[wsSource.Cells[9, 17], wsSource.Cells[NumberOfRows, 17]];
            wsImportToClear.Interior.Color = Color.White;

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
                        copyRow(wsMain, wsSource, rowS, 1, 18);

                        // Colors from https://safeery2k.wordpress.com/2013/06/19/system-drawing-knowncolor/

                        wsMain.Cells[rowS, 17].Value = DateTime.Now.ToString("s"); // put the short date time here
                        wsMain.Cells[rowS, 17].Interior.Color = Color.Cyan;
                       
                        wsSource.Cells[rowS, 17].Value = "updated";
                        wsSource.Cells[rowS, 17].Interior.Color = Color.Lime;
                    }
                }
            }

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

                }
            }
        }

    }


}
