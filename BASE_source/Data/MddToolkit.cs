using ExcelAlias = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System.Diagnostics;
using System.Web;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Office.Interop.Word;
using System.Windows.Media.Media3D;
using System.Runtime.InteropServices;

namespace BASE.Data
{
    [Serializable]
    public class MddToolkit : AbstractFile
    {

        const int cToolkitSearchUntilEmptyColumnNotIIGOV = 2;
        const int cToolkitEndTestNotIIGOV = 15; // test for end of file for brute find

        const int cToolkitSearchUntilEmptyColumnIIGOV = 3;
        const int cToolkitEndTestIIGOV = 3; // test for end of file for brute find

        const int cToolkitHeadingStartRow = 1;
        const int cToolkitMaxRows = 10000;

        // int numberOfProjects = 0;

        public List<PracticeArea_Element> MddToolkitPracticeAreas { get; set; }
            = new List<PracticeArea_Element>();
        public override bool LoadPersistantXMLdata()
        {
            try
            {
                // base.LoadPersistant(); override the base function, to load all information from here for this object and its parent
                if (File.Exists(_directoryFileNameXML))
                {
                    // If the directory and file name exists, laod the data
                    var xs = new XmlSerializer(typeof(MddToolkit));
                    using (FileStream xmlLoad = File.Open(_directoryFileNameXML, FileMode.Open))
                    {
                        var pData = (MddToolkit)xs.Deserialize(xmlLoad);
                        this.DirectoryFileName = pData._directoryFileName;

                        // *** Load the object elements belwo
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
            if (o is MddToolkit tko)
            {
                if (!Directory.Exists(Path.GetDirectoryName(_directoryFileNameXML)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(_directoryFileNameXML));
                }
                var xs = new XmlSerializer(typeof(MddToolkit));
                using (FileStream stream = File.Create(_directoryFileNameXML))
                {
                    xs.Serialize(stream, tko);
                }
            }
            else
            {
                throw new NotImplementedException("Object missmatched");

            }

        }

        public bool SetAllTo_FullyMet(System.Windows.Forms.Label lblStatus, out string resultMessage)
        {

            // *** Setup the main sheet
            // excelApp.Visible = true;

            // *** Load main
            //mainWorkbook = excelApp.Workbooks.Open(persistentData.OEdatabasePathFile);
            ExcelAlias.Workbook mainWorkbook;
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
            {
                resultMessage = "File not found, has it been moved or deleted?";
                return false;
            }
            string basePath = Path.GetDirectoryName(_directoryFileName);

            string statusStr = "Toolkit master:";
            lblStatus.Text = statusStr;

            int LastUsedRow = 1;
            foreach (ExcelAlias.Worksheet wksToolkitMaster in mainWorkbook.Worksheets)
            {
                bool IIAndGOV = false;
                bool ProcessPA = false;
                int columnToCheck = 3;
                switch (wksToolkitMaster.Name)
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
                        LastUsedRow = Helper.FindEndOfWorksheetBrute(wksToolkitMaster, cToolkitSearchUntilEmptyColumnNotIIGOV, cToolkitHeadingStartRow, cToolkitMaxRows, cToolkitEndTestNotIIGOV);
                        ProcessPA = true;
                        columnToCheck = 3;
                        IIAndGOV = false;
                        break;
                    case "GOV":
                    case "II":
                        // *** Find the number of rows
                        LastUsedRow = Helper.FindEndOfWorksheetBrute(wksToolkitMaster, cToolkitSearchUntilEmptyColumnIIGOV, cToolkitHeadingStartRow, cToolkitMaxRows, cToolkitEndTestIIGOV);
                        ProcessPA = true;
                        columnToCheck = 4;
                        IIAndGOV = true;
                        break;
                }

                if (ProcessPA == true)
                {
                    // *** Show the status
                    statusStr = statusStr + "[" + wksToolkitMaster.Name + $"{LastUsedRow}] ";
                    lblStatus.Text = statusStr;

                    // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                    //Range mainRange = wksToolkitMaster.Range["A" + cToolkitHeadingStartRow, "Z" + LastUsedRow];
                    ExcelAlias.Range mainRange = wksToolkitMaster.Range["A1", "Z" + LastUsedRow];

                    // *** Process the worksheet here
                    for (int row = cToolkitHeadingStartRow + 1; row <= LastUsedRow; row++)
                    {
                        // *** Find a DM, PM, LM or FM and set everything to 1,1,1
                        string cellValue = mainRange.Cells[row, columnToCheck]?.Value?.ToString() ?? "";
                        switch (cellValue.ToUpper())
                        {

                            case "DM":
                            case "PM":
                            case "LM":
                            case "FM":
                                mainRange.Cells[row, columnToCheck + 3].Value = 1;
                                mainRange.Cells[row, columnToCheck + 4].Value = 1;
                                mainRange.Cells[row, columnToCheck + 5].Value = 1;

                                break;

                        }

                        // *** See if a cell contains a PA and set the OU to FM
                        if (IIAndGOV)
                        {
                            string cellProcessValue = mainRange.Cells[row, columnToCheck - 2]?.Value?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(cellProcessValue))
                            {
                                // *** If a field contains a process, set the column to FM
                                mainRange.Cells[row, columnToCheck + 13].Value = "FM";
                            }
                        }
                        else
                        {

                            string cellPAvalue = mainRange.Cells[row, columnToCheck - 2]?.Value?.ToString() ?? "";
                            if (cellPAvalue.Length > wksToolkitMaster.Name.Length)
                            {
                                if (string.Compare(strA: cellPAvalue.Substring(0, wksToolkitMaster.Name.Length), strB: wksToolkitMaster.Name, culture: System.Globalization.CultureInfo.CurrentCulture,
                                    options: System.Globalization.CompareOptions.IgnoreCase) == 0)
                                {
                                    // *** strings are the same
                                    mainRange.Cells[row, columnToCheck + 13].Value = "FM";
                                }
                            }
                        }

                        // *** See if a cell  ontains NR and set it to S
                        string cellRatingValue = mainRange.Cells[row, columnToCheck + 15]?.Value?.ToString() ?? "";
                        if (cellRatingValue.ToUpper() == "NR") mainRange.Cells[row, columnToCheck + 15].Value = "S";
                    }
                }
            }

            mainWorkbook.Save();
            statusStr = statusStr + "done";
            lblStatus.Text = statusStr;

            //MessageBox.Show("Done");

            resultMessage = "All Full Met updated!";
            return true;
        }

        public bool PopulateToolkitFromOEdb(System.Windows.Forms.Label lblStatus,
            OEdbFile oeDbFile, CasPlanFile casPlanFile, out string resultMessage)
        {
            resultMessage = "Successfull.";

            //mainWorkbook = excelApp.Workbooks.Open(persistentData.OEdatabasePathFile);
            ExcelAlias.Workbook mddToolkitWorkbook;
            if ((mddToolkitWorkbook = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
            {
                resultMessage = "MDD Toolkit file not found, has it been moved or deleted?";
                return false;
            }
            string basePath = Path.GetDirectoryName(_directoryFileName);

            ExcelAlias.Workbook oeDbWorkbook;
            if ((oeDbWorkbook = Helper.CheckIfOpenAndOpenXlsx(oeDbFile._directoryFileName)) == null)
            {
                resultMessage = "OEdb file not found, has it been moved or deleted?";
                return false;
            }

            ExcelAlias.Workbook casPlanWorkbook;
            if ((casPlanWorkbook = Helper.CheckIfOpenAndOpenXlsx(casPlanFile._directoryFileName)) == null)
            {
                resultMessage = "CAS Plan file not found, has it been moved or deleted?";
                return false;
            }

            // *** Update OEdbFile record
            // * Build data for MDDToolKit
            lblStatus.Text = "";
            foreach (ExcelAlias.Worksheet oeWksSource in oeDbWorkbook.Sheets)
            {
                // Clear filters if it is set
                // https://stackoverflow.com/questions/13204064/turn-off-filters
                if (oeWksSource.AutoFilter != null)
                {
                    oeWksSource.AutoFilterMode = false;
                }

                string paCodeAsString = string.Empty;

                switch (oeWksSource.Name)
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
                        PracticeArea_Element aPracticeArea_non_II_GOV = new PracticeArea_Element();
                        aPracticeArea_non_II_GOV.AcronymName = oeWksSource.Name;

                        if (Update_OEdb_WithWorksheetsStats_non_II_GOV(oeWksSource, ref aPracticeArea_non_II_GOV) == true)
                        {
                            MddToolkitPracticeAreas.Add(aPracticeArea_non_II_GOV);
                        };
                        lblStatus.Text = lblStatus.Text + " " + oeWksSource.Name;
                        break;
                    case "II":
                    case "GOV":
                        PracticeArea_Element aPracticeArea_II_GOV = new PracticeArea_Element();
                        aPracticeArea_II_GOV.AcronymName = oeWksSource.Name;

                        if (Extract_OEdb_From_Wks_for_II_GOV2(oeWksSource, ref aPracticeArea_II_GOV) == true)
                        {
                            MddToolkitPracticeAreas.Add(aPracticeArea_II_GOV);
                        };
                        lblStatus.Text = lblStatus.Text + " " + oeWksSource.Name;
                        break;

                }

            }

            // * Use data to populat MDDToolkit
            ExcelAlias.Worksheet worksheet2;
            ExcelAlias.Range range2;


            int endRow;
            int startRowX;
            int endRowX;
            lblStatus.Text = "";
            int numberOfProjects = casPlanFile.WorkUnitList2.Count;
            foreach (PracticeArea_Element aPA2 in MddToolkitPracticeAreas)
            {

                lblStatus.Text = lblStatus.Text + aPA2.AcronymName + " ";
                worksheet2 = mddToolkitWorkbook.Worksheets[aPA2.AcronymName];

                switch (aPA2.AcronymName)
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

                        startRowX = 1; // start row of a MDD toolkit worksheet
                        endRow = Helper.FindEndOfWorksheet(worksheet2, 2, startRowX, 5000);
                        if (endRow != -1)
                        {

                            if (worksheet2 != null)
                            {
                                foreach (Practice_Element aPractice2 in aPA2.practice_Elements)
                                {

                                    if (FindTheStartAndEnd_Practice(worksheet2,
                                        aPractice2.CodeAndNumber, startRowX, endRow,
                                        out startRowX, out endRowX))
                                    {
                                        string projectChar;
                                        string projectName;
                                        worksheet2.Cells[startRowX, 15].Value = aPractice2.Char;
                                        
                                        worksheet2.Cells[startRowX, 17].Value = "S"; // assign rating 'S'


                                        for (int rowx2 = startRowX; rowx2 < startRowX + numberOfProjects; rowx2++)

                                        //                                            for (int rowx2 = startRowX; rowx2 < startRowX + numberOfProjects - 1; rowx2++)
                                        {
                                            // find a project that is not OoS
                                            projectChar = worksheet2.Cells[rowx2, 3]?.Value2?.ToString();
                                            if (projectChar?.ToLower() != "OoS".ToLower())
                                            {
                                                projectName = worksheet2.Cells[rowx2, 2]?.Value2?.ToString();
                                                projectName = projectName?.ToLower();

                                                var workUnit = aPractice2?.PrjSup_Elements?
                                                    .FirstOrDefault(w => w.projectSupportName.ToLower() == projectName);
                                                if (workUnit != null)
                                                {

                                                    // Console.WriteLine($"Found WorkUnit: {workUnit.Name}");
                                                    worksheet2.Cells[rowx2, 6].Value = 1;  // Assigning a number
                                                    worksheet2.Cells[rowx2, 7].Value = 1;  // Assigning a number
                                                                                           //worksheet2.Cells[rowx2, 8].Value = workUnit.TheOECount;  // Assigning a number
                                                    worksheet2.Cells[rowx2, 8].Value = // Column I OE Count
                                                        workUnit.oE_Elements.Count;
                                                    
                                                    worksheet2.Cells[rowx2, 10].Value = 2;
                                                    worksheet2.Cells[rowx2, 11].Value =
                                                        string.Join(", ", aPractice2.SessionList);
                                                    worksheet2.Cells[rowx2, 12].Value =
                                                        string.Join(", ", aPractice2.ParticipantList);

                                                    worksheet2.Cells[rowx2, 14].Value = // n
                                                        BuildStrengthRecommendation(aPractice2.StrengthList, 
                                                        aPractice2.RecommendationList);

                                                    string result = BuildWeaknessList(aPractice2.WeaknessesList);
                                                    worksheet2.Cells[rowx2, 9].Value = result; // i
                                                    worksheet2.Cells[rowx2, 16].Value = result; // p

                                                    // worksheet2.Cells[rowx2, 3].Style.Numberformat.Format = "0"; // Ensuring it's formatted as a number
                                                }
                                                else
                                                {
                                                    // Console.WriteLine("WorkUnit not found.");
                                                }

                                            }

                                        }

                                    }

                                }
                            }
                        }
                        break;

                    case "II":
                    case "GOV":

                        int currentMianRow = 2; // start row of a MDD toolkit worksheet
                        endRow = Helper.FindEndOfWorksheet(worksheet2, 3, currentMianRow, 5000);
                        if (endRow != -1)
                        {

                            if (worksheet2 != null)
                            {
                                int NumberOfPractices = aPA2.practice_Elements.Count();
                                foreach (Practice_Element aPractice2 in aPA2.practice_Elements)
                                {
                                    worksheet2.Cells[currentMianRow, 18].Value = "S"; // assign rating 'S'

                                    int NumberOfProcessess = aPractice2.ProcessElements.Count();
                                    foreach (Process_Element aProcess2 in aPractice2.ProcessElements)
                                    {
                                        // * At this point, the process should be defined at startRow,2 (if not break error)
                                        string theProcess = worksheet2?.Cells[currentMianRow, 2]?.Value?.ToString();
                                        if (theProcess != aProcess2.ProcessName)
                                        {
                                            MessageBox.Show($"{aPA2.AcronymName}:{aPractice2.CodeAndNumber} Process missmatch ({aProcess2.ProcessName} vs {theProcess})");
                                            Debug.WriteLine($"Practice area:{aPA2.AcronymName} Practice:{aPractice2.CodeAndNumber} Process:{aProcess2.ProcessName}");
                                            return false;
                                        }
                                        // *** Else set OU char
                                        worksheet2.Cells[currentMianRow, 16].Value = aProcess2.Char;

                                        // *** Cont look at what was read into aPA2/aPractice2/aProcess2,
                                        // *** Add the full number of projects, becuase that is how toolkit is setup
                                        int ProjectItems = 20; // Standard list by ISACA
                                        int NumberOfProjectSupportItems = casPlanFile.WorkUnitList2.Count;

                                        int tempRow = currentMianRow;

                                        for (int WorkUnitListCount = 1; WorkUnitListCount <= NumberOfProjectSupportItems; WorkUnitListCount++)
                                        {
                                            // ** Step 1: Read the first one
                                            string projectName = worksheet2?.Cells[tempRow, 3]?.Value?.ToString();
                                            if (!string.IsNullOrEmpty(projectName))
                                            {
                                                var prjSupFoundInList = aProcess2.PrjSup_Elements.FirstOrDefault(
                                                    x => x.projectSupportName.Equals(projectName, StringComparison.OrdinalIgnoreCase));
                                                if (prjSupFoundInList != null)
                                                {
                                                    // ** Step 2a: Find it in aProcess2.PrjSup_Elements (if found, update)
                                                    worksheet2.Cells[tempRow, 7].Value = 1; // Column G Sufficnet
                                                    worksheet2.Cells[tempRow, 8].Value = 1; // Column H Affirmatino
                                                    worksheet2.Cells[tempRow, 9].Value = // Column I OE Count
                                                        prjSupFoundInList.oE_Elements.Count;
                                                    worksheet2.Cells[tempRow, 10].Value = // Column J
                                                        prjSupFoundInList.weaknessStr;
                                                    worksheet2.Cells[tempRow, 11].Value = 2; // Column K Readiness Revie Totals
                                                    worksheet2.Cells[tempRow, 12].Value = // Column L
                                                        prjSupFoundInList.sessionName;
                                                    worksheet2.Cells[tempRow, 13].Value = // Column M
                                                        string.Join(", ", prjSupFoundInList?.participantList) ?? "";

                                                    worksheet2.Cells[tempRow, 15].Value = // O
                                                            BuildStrengthRecommendation(aPractice2.StrengthList, 
                                                            aPractice2.RecommendationList);

                                                    string result = BuildWeaknessList(aPractice2.WeaknessesList);
                                                    worksheet2.Cells[tempRow, 10].Value = result; // J
                                                    worksheet2.Cells[tempRow, 17].Value = result; // Q

                                                }
                                                else
                                                {
                                                    // ** Step 2b: If NOT found, ignore

                                                }
                                            }
                                            tempRow++;
                                        }
                                        currentMianRow = currentMianRow + 20; // *** ISACA allocated project support items.

                                    }
                                    currentMianRow = currentMianRow + 20 * (25 - NumberOfProcessess); // Offset ISAC count
                                }
                            }
                        }
                        break;
                }


            }

            lblStatus.Text = lblStatus.Text + "done";
            resultMessage = "Successfull completed";
            return true;

        }

        string BuildStrengthRecommendation(List<string> strengthList1, List<string> recommendationList1)
        {
            string strengthListCombined;
            if (strengthList1.Count == 0)
            {
                strengthListCombined = string.Empty;
            }
            else
            {
                strengthListCombined = string.Join(", ", strengthList1) + " (S)\n";
            }
            string recommendationListCombined;
            if (strengthList1.Count == 0)
            {
                recommendationListCombined = string.Empty;
            }
            else
            {
                recommendationListCombined = string.Join(", ", strengthList1) + " (R)";
            }
            return strengthListCombined + recommendationListCombined;
        }

        string BuildWeaknessList(List<string> weaknessList1)
        {
            if (weaknessList1.Count == 0)
            {
                return string.Empty;
            }
            else
            {
                return string.Join(", ", weaknessList1) + " (W)";
            }

        }



        public bool SetupOUandProcessScope(System.Windows.Forms.Label lblStatus,
            //  MddToolkit mddToolkitFile,
            CasPlanFile casPlanFile,
            out string resultMessage)
        {

            resultMessage = "Successfull.";

            // *** Open the workbook
            ExcelAlias.Workbook mddToolkitWorkbook;
            if ((mddToolkitWorkbook = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
            {
                resultMessage = "MDD Toolkit file not found, has it been moved or deleted?";
                return false;
            }

            // *** Build the "Organizaional Unit Summary"
            ExcelAlias.Worksheet ouSummaryWks = mddToolkitWorkbook.Worksheets["Organizational Unit Summary"];
            ouSummaryWks.Cells[17, 2].Value2 = casPlanFile.Organisation.Name;
            ouSummaryWks.Cells[18, 2].Value2 = casPlanFile.OrganizationalUnit.Name;
            string fullAddess = casPlanFile.Organisation.FullAddress;

            ouSummaryWks.Cells[19, 2].Value2 = casPlanFile.Organisation.FullAddress;
            ouSummaryWks.Cells[20, 2].Value2 = casPlanFile.Phase2Start;
            ouSummaryWks.Cells[21, 2].Value2 = casPlanFile.Phase2End;


            // Set Target Level
            switch (casPlanFile.OrganizationalUnit.MaturityLevel)
            {
                case 2:
                    ouSummaryWks.Cells[27, 2].Value2 = "Level 2 (Managed)";
                    break;
                case 3:
                    ouSummaryWks.Cells[27, 2].Value2 = "Level 3 (Defined)";
                    break;
                case 4:
                    ouSummaryWks.Cells[27, 2].Value2 = "Level 4 (Quantitatively Managed)";
                    break;
                case 5:
                    ouSummaryWks.Cells[27, 2].Value2 = "Level 5 (Optimizing)";
                    break;
                default:
                    ouSummaryWks.Cells[27, 2].Value2 = "Level 1 (Initial)";
                    break;
            }

            // Set DEV = Yes
            ouSummaryWks.Cells[29, 2].Value2 = "Yes";
            

            // List the work units
            // numberOfProjects = casPlanFile.WorkUnitList2.Count();
            ouSummaryWks.Cells[38, 2].Value2 = casPlanFile.WorkUnitList2.Count();
            int workUnitRow = 40;
            foreach (WorkUnit aWorkUnit in casPlanFile.WorkUnitList2)
            {
                ouSummaryWks.Cells[workUnitRow, 2].Value2 = aWorkUnit.Name;
                workUnitRow++;
            }


            // *** Open the worksheet
            ExcelAlias.Worksheet toolkitProcess = mddToolkitWorkbook.Worksheets["Processes"];
            List<OUProcess> ouProcessList = casPlanFile.OUProcessesList2;
            if (ToolkitSelectScopeForProceses(toolkitProcess, ouProcessList, casPlanFile.WorkUnitList2.Count()))
            {
            }
            else
            {
                resultMessage = "Error. Could not populate scope for process and support functions.";
                return false;
            }
            // *** Open the worksheet
            ExcelAlias.Worksheet toolkitOUS = mddToolkitWorkbook.Worksheets["Organizational Unit Summary"];
            List<WorkUnit> workUnitList2 = casPlanFile.WorkUnitList2;
            if (ToolkitSelectScopeForProjects(toolkitOUS, workUnitList2))
            {
            }
            else
            {
                resultMessage = "Error. Could not populate scope for projects and support functions.";
                return false;
            }



            resultMessage = "All good";
            return true;
        }

        bool ToolkitSelectScopeForProceses(ExcelAlias.Worksheet toolkitProcess,
            List<OUProcess> processList2, int workUnitCount)
        {

            int numberOfProcesess = processList2.Count();
            if (numberOfProcesess == 0) return false;

            toolkitProcess.Cells[1, 3].Value = numberOfProcesess;

            int row = 3;
            processList2 = processList2.OrderBy(x => x.Name).ToList();
            foreach (OUProcess aProcess in processList2)
            {
                toolkitProcess.Cells[row, 2].Value = aProcess.Name;
                // check accross projects and list
                for (int col = 4; col <= workUnitCount + 4 - 1; col++)
                {
                    string lisedWorkUnit = toolkitProcess?.Cells[2, col]?.Value?.ToString();
                    // *** Is this listedWorkUnit in the aProcess.WorkUnit
                    var WorkUnitListed = aProcess.WorkUnits.FirstOrDefault(x => x.Name.ToLower() == lisedWorkUnit.ToLower());
                    if (WorkUnitListed != null)
                    {
                        // It is an IS
                        toolkitProcess.Cells[row, col].Value = "IS";
                    }

                }
                row++;

            }


            return true;

        }

        bool ToolkitSelectScopeForProjects(ExcelAlias.Worksheet toolkitOUS2,
            List<WorkUnit> workUnitList2)
        {
            int workUnitCount = workUnitList2.Count;
            int paColumnList = 6;
            int projectRow = 9;
            string workUnitSelectedStr;
            string PAselectedStr;
            WorkUnit WorkUnitX;

            for (int column = paColumnList + 1; column <= workUnitCount + paColumnList; column++)
            {
                workUnitSelectedStr = toolkitOUS2?.Cells[projectRow, column]?.Value?.ToString();
                WorkUnitX = MatchingWorkUnit(workUnitSelectedStr, workUnitList2);

                if (WorkUnitX != null)
                {
                    for (int row = projectRow + 1; row <= 48; row++)
                    {
                        PAselectedStr = toolkitOUS2?.Cells[row, paColumnList]?.Value?.ToString();
                        if (!string.IsNullOrEmpty(PAselectedStr))
                        {
                            if (Enum.TryParse<EPAcode>(PAselectedStr, true, out EPAcode EPAcodeResult))
                            {
                                //var PAx = WorkUnitX.PAlist.FirstOrDefault(x => x.Name == PAselectedStr);
                                var PAx = WorkUnitX.PAlist.FirstOrDefault(x =>
                                    x.PAcode == EPAcodeResult);
                                if (PAx != null)
                                {
                                    toolkitOUS2.Cells[row, column].Value = "IS";
                                }
                            }
                        }

                    }
                }
            }


            return true;
        }
        WorkUnit MatchingWorkUnit(string workUnitSelected, List<WorkUnit> workUnitList2)
        {
            if (workUnitList2 == null) return null;

            foreach (WorkUnit aWorkUnit in workUnitList2)
            {
                if (aWorkUnit.Name.Length >= 6)
                {
                    if (aWorkUnit.Name.Substring(0, 6).ToLower() == workUnitSelected.ToLower())
                        return aWorkUnit;
                }
                else
                {
                    if (aWorkUnit.Name.ToLower() == workUnitSelected.ToLower())
                        return aWorkUnit;
                }
            }
            return null;

        }

        bool Update_OEdb_WithWorksheetsStats_non_II_GOV(ExcelAlias.Worksheet wsSource2, ref PracticeArea_Element aPracticeArea2)
        {

            // *** Find the number of rows in wsSource
            int NumberOfRows2 = Helper.FindEndOfWorksheet(wsSource2, OEdbFile.cDemixOEToolSearchUntilEmptyColumn,
                    OEdbFile.cDemixOEToolHeadingStartRow, OEdbFile.cDemixOEToolMaxRows);

            // aPracticeArea = new PracticeArea();

            aPracticeArea2 = new PracticeArea_Element();

            aPracticeArea2.AcronymName = wsSource2?.Name;
            Practice_Element aPractice = new Practice_Element();
            PrjSup_Element aPrjSup = new PrjSup_Element();
            OE_Element aOEelement = new OE_Element();
            Regex regex1 = new Regex(@"(\[\d+\]\s*.*?)(?:\((.*?)\))");

            for (int row = NumberOfRows2; row >= 8; row--)
            {
                string headingStr = wsSource2?.Cells[row, 1]?.Value?.ToString();

                switch (headingStr?.ToLower())
                {
                    case "1 prac_group":
                        // ignore
                        break;
                    case "2 prac_ou":
                        aPractice.CodeAndNumber = wsSource2?.Cells[row, 2]?.Value?.ToString();

                        //aPractice.WeaknessesList = wsSource2?.Cells[row, 12]?.Value?.ToString(); // is build by 4 Prac_instan
                        //aPractice.StrengthList = wsSource2?.Cells[row, 13]?.Value?.ToString(); // is build by 4 Prac_instan
                        //aPractice.RecommendationList = wsSource2?.Cells[row, 14]?.Value?.ToString(); // is build by 4 Prac_instan
                        aPractice.Char = wsSource2?.Cells[row, 15]?.Value?.ToString();

                        aPracticeArea2.practice_Elements.Add(aPractice);
                        aPractice = new Practice_Element();
                        break;
                    case "4 prac_instan":
                        aPrjSup.projectSupportName = wsSource2?.Cells[row, 3]?.Value?.ToString().Trim();

                        // Regular expression to capture the required parts
                        string sessionParticipant = wsSource2?.Cells[row, 8]?.Value?.ToString().Trim();
                        Match match = regex1.Match(sessionParticipant);
                        aPrjSup.sessionName = match.Groups[1].Value.Trim();  // "[4] Engineering"
                        string participantStr = match.Groups[2].Value.Trim();
                        aPrjSup.participantList = OEdbProcessors.GetParticipants(participantStr);

                        aPrjSup.weaknessStr = wsSource2?.Cells[row, 12]?.Value?.ToString().Trim();
                        aPrjSup.strengthStr = wsSource2?.Cells[row, 13]?.Value?.ToString().Trim();
                        aPrjSup.recommendationStr = wsSource2?.Cells[row, 14]?.Value?.ToString().Trim();

                        aPrjSup.Char = wsSource2?.Cells[row, 15]?.Value?.ToString().Trim();

                        aPractice.PrjSup_Elements.Add(aPrjSup);
                        aPrjSup = new PrjSup_Element();
                        break;
                    case "5 oe":
                        aOEelement.ProjectName = wsSource2?.Cells[row, 2]?.Value?.ToString();
                        string OEtypeStr = wsSource2?.Cells[row, 6]?.Value?.ToString();
                        switch (OEtypeStr.ToLower())
                        {
                            case "ok file":
                                aOEelement.OeStatus = E_OEStatus.OkFile;
                                break;
                            case "ok directory":
                                aOEelement.OeStatus = E_OEStatus.OkDirectory;
                                break;
                            case "not ok":
                                aOEelement.OeStatus = E_OEStatus.NotOk;

                                break;
                            default:
                                aOEelement.OeStatus = E_OEStatus.None;

                                break;
                        }
                        string SuffStr = wsSource2?.Cells[row, 9]?.Value?.ToString();
                        aOEelement.Sufficient = SuffStr?.ToLower().Trim() == "yes" ? E_YesNo.Yes : E_YesNo.No;

                        aPrjSup.oE_Elements.Add(aOEelement);
                        aOEelement = new OE_Element();
                        break;

                    default:
                        break;
                }
            }

            aPracticeArea2.practice_Elements.Sort((x, y) => string.Compare(x.CodeAndNumber, y.CodeAndNumber, true));

            // summarise participants from 4_Prac_Instan into 2 Prac_OU
            // summarise sessions from 4_Prac_instan into 2 Prac_OU
            foreach (Practice_Element pe in aPracticeArea2.practice_Elements) // 2 Prac_OU
            {

                foreach (PrjSup_Element pse in pe.PrjSup_Elements) // 4 Prac_Instan
                {
                    pe.UpdateParticipants(pse.participantList);
                    pe.UpdateSessions(pse.sessionName);
                    pe.UpdateStrengths(pse.strengthStr);
                    pe.UpdateWeaknesses(pse.weaknessStr);
                    pe.UpdateRecommendations(pse.recommendationStr);

                    int oeCount = 0;
                    foreach (OE_Element oe in pse.oE_Elements) //5 OE
                    {
                        oeCount = oeCount + oe.Sufficient == E_YesNo.Yes ? 1 : 0;
                    }
                }
            }


            return true;

        }

        bool Extract_OEdb_From_Wks_for_II_GOV2(ExcelAlias.Worksheet wsSource2, ref PracticeArea_Element aPracticeArea2)
        {

            // *** Find the number of rows in wsSource
            int NumberOfRows2 = Helper.FindEndOfWorksheet(wsSource2, OEdbFile.cDemixOEToolSearchUntilEmptyColumn,
                    OEdbFile.cDemixOEToolHeadingStartRow, OEdbFile.cDemixOEToolMaxRows);

            // aPracticeArea = new PracticeArea();

            aPracticeArea2 = new PracticeArea_Element();

            aPracticeArea2.AcronymName = wsSource2?.Name;
            Process_Element aProcess_Element = new Process_Element();
            Practice_Element aPractice = new Practice_Element();
            PrjSup_Element aPrjSub_Element = new PrjSup_Element();
            OE_Element aOEelement = new OE_Element();
            Regex regex1 = new Regex(@"(\[\d+\]\s*.*?)(?:\((.*?)\))");

            for (int row = NumberOfRows2; row >= 8; row--)
            {
                string headingStr = wsSource2?.Cells[row, 1]?.Value?.ToString();

                switch (headingStr?.ToLower())
                {
                    case "1 prac_group":
                        // ignore
                        break;
                    case "2 prac_ou":
                        aPractice.CodeAndNumber = wsSource2?.Cells[row, 2]?.Value?.ToString();
                        aPractice.Char = wsSource2?.Cells[row, 15]?.Value?.ToString();
                        aPractice.ProcessElements = aPractice.ProcessElements.OrderBy(x => x.ProcessName).ToList();

                        aPracticeArea2.practice_Elements.Add(aPractice);

                        aPractice = new Practice_Element();
                        break;
                    case "3 process":
                        aProcess_Element.CodeAndNumber = wsSource2?.Cells[row, 2]?.Value?.ToString();
                        aProcess_Element.ProcessName = wsSource2?.Cells[row, 3]?.Value?.ToString();
                        aProcess_Element.Char = wsSource2?.Cells[row, 15]?.Value?.ToString();

                        aPractice.ProcessElements.Add(aProcess_Element);
                        aProcess_Element = new Process_Element();
                        break;
                    case "4 prac_instan":
                        aPrjSub_Element.projectSupportName = wsSource2?.Cells[row, 3]?.Value?.ToString().Trim();

                        // Regular expression to capture the required parts
                        string sessionParticipant = wsSource2?.Cells[row, 8]?.Value?.ToString().Trim();
                        Match match = regex1.Match(sessionParticipant);
                        aPrjSub_Element.sessionName = match.Groups[1].Value.Trim();  // "[4] Engineering"
                        string participantStr = match.Groups[2].Value.Trim();
                        aPrjSub_Element.participantList = OEdbProcessors.GetParticipants(participantStr);

                        aPrjSub_Element.weaknessStr = wsSource2?.Cells[row, 12]?.Value?.ToString().Trim();
                        aPrjSub_Element.strengthStr = wsSource2?.Cells[row, 13]?.Value?.ToString().Trim();
                        aPrjSub_Element.recommendationStr = wsSource2?.Cells[row, 14]?.Value?.ToString().Trim();

                        aPrjSub_Element.Char = wsSource2?.Cells[row, 15]?.Value?.ToString().Trim();

                        aProcess_Element.PrjSup_Elements.Add(aPrjSub_Element); // For II and GOV map to "3 Process" not "2 Prac_OU"
                        aPrjSub_Element = new PrjSup_Element();
                        break;
                    case "5 oe":
                        aOEelement.ProjectName = wsSource2?.Cells[row, 2]?.Value?.ToString();
                        string OEtypeStr = wsSource2?.Cells[row, 6]?.Value?.ToString();
                        switch (OEtypeStr.ToLower())
                        {
                            case "ok file":
                                aOEelement.OeStatus = E_OEStatus.OkFile;
                                break;
                            case "ok directory":
                                aOEelement.OeStatus = E_OEStatus.OkDirectory;
                                break;
                            case "not ok":
                                aOEelement.OeStatus = E_OEStatus.NotOk;

                                break;
                            default:
                                aOEelement.OeStatus = E_OEStatus.None;

                                break;
                        }
                        string SuffStr = wsSource2?.Cells[row, 9]?.Value?.ToString();
                        aOEelement.Sufficient = SuffStr?.ToLower().Trim() == "yes" ? E_YesNo.Yes : E_YesNo.No;

                        aPrjSub_Element.oE_Elements.Add(aOEelement);
                        aOEelement = new OE_Element();
                        break;

                    default:
                        break;
                }
            }

            aPracticeArea2.practice_Elements.Sort((x, y) => string.Compare(x.CodeAndNumber, y.CodeAndNumber, true));

            // summarise participants from 4_Prac_Instan into 2 Prac_OU
            // summarise sessions from 4_Prac_instan into 2 Prac_OU
            foreach (Practice_Element pe in aPracticeArea2.practice_Elements) // 2 Prac_OU
            {
                foreach (Process_Element processElement in pe.ProcessElements)
                {

                    foreach (PrjSup_Element pse in processElement.PrjSup_Elements) // 4 Prac_Instan
                    {

                        processElement.UpdateParticipants(pse.participantList);
                        processElement.UpdateSessions(pse.sessionName);
                        processElement.UpdateStrengths(pse.strengthStr);
                        processElement.UpdateWeaknesses(pse.weaknessStr);
                        processElement.UpdateRecommendations(pse.recommendationStr);


                        int oeCount = 0;
                        foreach (OE_Element oe in pse.oE_Elements) //5 OE
                        {
                            oeCount = oeCount + oe.Sufficient == E_YesNo.Yes ? 1 : 0;
                        }
                    }
                    pe.UpdateParticipants(processElement.ParticipantList);
                    pe.UpdateSessions(processElement.SessionList);
                    pe.UpdateStrengths(processElement.StrengthList);
                    pe.UpdateWeaknesses(processElement.WeaknessesList);
                    pe.UpdateRecommendations(processElement.RecommendationList);

                }
            }


            return true;

        }


        bool FindTheStartAndEnd_Practice(ExcelAlias.Worksheet wks3, string practice3,
           int startRow, int endRow,
           out int startRowPractice, out int endRowPractice)
        {
            // startRow containst the location to start searchign for practice3.
            startRowPractice = -1;
            endRowPractice = -1;
            int searchRow = startRow;
            string cellStr = string.Empty;

            // Find the first occurance
            for (searchRow = startRow; searchRow <= endRow; searchRow++)
            {
                cellStr = wks3.Cells[searchRow, 1]?.Value?.ToString();
                if (cellStr?.ToLower() == practice3?.ToLower())
                {
                    // *** First occurance found
                    startRowPractice = searchRow;
                    break;
                }
            }
            if (startRowPractice == -1)
            {
                // Could not find the practice3
                return false;
            }
            // Found the practice3, for example CAR 1.1, now search the end
            for (searchRow = startRowPractice; searchRow <= endRow; searchRow++)
            {
                cellStr = wks3.Cells[searchRow, 1]?.Value?.ToString();
                if (cellStr.ToLower() == practice3.ToLower())
                {
                    // *** Second occurance found, not written to start where previous end, so 
                    // startRowPractice could be the same as 
                    endRowPractice = searchRow;
                    break;
                }
                else
                {
                    break;
                }
            }
            return true;
        }
    }
}