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

        int numberOfProjects = 0;

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

            // write status here string statusStr = "Toolkit master:";
            // write status herelblStatus.Text = statusStr;

            // *** Build the "Organizaional Unit Summary"
            ExcelAlias.Worksheet ouSummaryWks = mddToolkitWorkbook.Worksheets["Organizational Unit Summary"];
            ouSummaryWks.Cells[17, 2].Value2 = casPlanFile.Organisation.Name;
            ouSummaryWks.Cells[18, 2].Value2 = casPlanFile.OrganizationalUnit.Name;
            string fullAddess = casPlanFile.Organisation.FullAddress;

            ouSummaryWks.Cells[19, 2].Value2 = casPlanFile.Organisation.FullAddress;
            ouSummaryWks.Cells[20, 2].Value2 = casPlanFile.Phase2Start;
            ouSummaryWks.Cells[21, 2].Value2 = casPlanFile.Phase2End;

            // List the work units
            int workUnitRow = 40;
            foreach (WorkUnit aWorkUnit in casPlanFile.WorkUnitList2)
            {
                ouSummaryWks.Cells[workUnitRow, 2].Value2 = aWorkUnit.Name;
                workUnitRow++;
            }
            ouSummaryWks.Cells[38, 2].Value2 = casPlanFile.WorkUnitList2.Count();
            numberOfProjects = casPlanFile.WorkUnitList2.Count();


            // *** Update OEdbFile record

            // * Build data for MDDToolKit
            lblStatus.Text = "";
            foreach (ExcelAlias.Worksheet wsSource in oeDbWorkbook.Sheets)
            {

                // Clear filters if it is set
                // https://stackoverflow.com/questions/13204064/turn-off-filters
                if (wsSource.AutoFilter != null)
                {
                    wsSource.AutoFilterMode = false;
                }


                string paCodeAsString = string.Empty;

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
                        //    case "II":
                        //   case "GOV":

                        // *** extract the source and destination range https://stackoverflow.com/questions/910400/reading-from-excel-range-into-multidimensional-array-c-sharp
                        //ProcessRowsUsingObject(wsMain, wsSource, NumberOfRows, ref statusStr);
                        PracticeArea_Element aPracticeArea = new PracticeArea_Element();
                        aPracticeArea.AcronymName = wsSource.Name;

                        if (Update_OEdb_WithWorksheetsStats2(wsSource, ref aPracticeArea) == true)
                        {
                            MddToolkitPracticeAreas.Add(aPracticeArea);
                        };
                        lblStatus.Text = lblStatus.Text + " " + wsSource.Name;
                        break;
                }

            }

            // * Use data to populat MDDToolkit
            ExcelAlias.Worksheet worksheet2;
            ExcelAlias.Range range2;

            int startRow;
            int endRow;
            int startRowX;
            int endRowX;
            foreach (PracticeArea_Element aPA2 in MddToolkitPracticeAreas)
            {
                worksheet2 = mddToolkitWorkbook.Worksheets[aPA2.AcronymName];
                startRow = 1; // start row of a MDD toolkit worksheet
                endRow = Helper.FindEndOfWorksheet(worksheet2, 2, startRow, 5000);
                if (endRow != -1)
                {

                    if (worksheet2 != null)
                    {
                        foreach (Practice_Element aPractice2 in aPA2.practice_Elements)
                        {

                            if (FindTheStartAndEnd_Practice(worksheet2,
                                aPractice2.CodeAndNumber, startRow, endRow,
                                out startRowX, out endRowX))
                            {
                                string projectChar;
                                string projectName;
                                worksheet2.Cells[startRowX, 15].Value = aPractice2.CharStr;
                                // worksheet2.Cells[startRowX, 17].Value = "S";


                                for (int rowx2 = startRowX; rowx2 < startRowX + numberOfProjects - 1; rowx2++)
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
                                            worksheet2.Cells[rowx2, 9].Value = string.Join(", ", aPractice2.WeaknessesList);  // Assigning a number
                                            worksheet2.Cells[rowx2, 10].Value = 2;
                                            worksheet2.Cells[rowx2, 11].Value =
                                                string.Join(", ", aPractice2.SessionList);
                                            worksheet2.Cells[rowx2, 12].Value =
                                                string.Join(", ", aPractice2.ParticipantList);
                                            worksheet2.Cells[rowx2, 14].Value =
                                                string.Join(", ", aPractice2.StrengthList) + "(S)\n " +
                                                string.Join(", ", aPractice2.RecommendationList) + "(R)\n ";
                                            worksheet2.Cells[rowx2, 17].Value =
                                                string.Join(", ", aPractice2.WeaknessesList);

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
            }

            lblStatus.Text = lblStatus.Text + "done";
            resultMessage = "Successfull completed";
            return true;

        }

        public bool SetupOUandProcessScope(System.Windows.Forms.Label lblStatus,
            //  MddToolkit mddToolkitFile,
            CasPlanFile casPlanFile,
            out string resultMessage)
        {

            resultMessage = "Successfull.";

            //mainWorkbook = excelApp.Workbooks.Open(persistentData.OEdatabasePathFile);
            ExcelAlias.Workbook mddToolkitWorkbook;
            if ((mddToolkitWorkbook = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
            {
                resultMessage = "MDD Toolkit file not found, has it been moved or deleted?";
                return false;
            }

            // Use OUProcessList2 to populate Tab: Organizational Unit Summary
            // Row 9 Colum F (6) is the place where the Scope needs to be updated

            // go through the projects (left to right)
            // for each go through the list of PAs Top to bottom F10 to F51
            // if you find the project in the OUProcessList and you find the PA, then marke it
            ExcelAlias.Worksheet toolkitOUS = mddToolkitWorkbook.Worksheets["Organizational Unit Summary"];

            List<OUProcess> ouProcessList = casPlanFile.OUProcessesList2;
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

        bool ToolkitSelectScopeForProceses()
        {

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
                } else
                {
                    if (aWorkUnit.Name.ToLower() == workUnitSelected.ToLower())
                        return aWorkUnit;
                }
            }
            return null;

        }


        bool Update_OEdb_WithWorksheetsStats2(ExcelAlias.Worksheet wsSource2, ref PracticeArea_Element aPracticeArea2)
        {

            // *** Find the number of rows in wsSource
            int NumberOfRows2 = Helper.FindEndOfWorksheet(wsSource2, OEdbFile.cDemixOEToolSearchUntilEmptyColumn,
                    OEdbFile.cDemixOEToolHeadingStartRow, OEdbFile.cDemixOEToolMaxRows);

            // aPracticeArea = new PracticeArea();

            aPracticeArea2 = new PracticeArea_Element();

            aPracticeArea2.AcronymName = wsSource2?.Name;
            Practice_Element aPractice = new Practice_Element();
            PrjSup_Element aWorkUnit = new PrjSup_Element();
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
                        aPractice.CharStr = wsSource2?.Cells[row, 15]?.Value?.ToString();

                        aPracticeArea2.practice_Elements.Add(aPractice);
                        aPractice = new Practice_Element();
                        break;
                    case "4 prac_instan":
                        aWorkUnit.projectSupportName = wsSource2?.Cells[row, 3]?.Value?.ToString().Trim();

                        // Regular expression to capture the required parts
                        string sessionParticipant = wsSource2?.Cells[row, 8]?.Value?.ToString().Trim();
                        Match match = regex1.Match(sessionParticipant);
                        aWorkUnit.sessionName = match.Groups[1].Value.Trim();  // "[4] Engineering"
                        string participantStr = match.Groups[2].Value.Trim();
                        aWorkUnit.participantList = OEdbProcessors.GetParticipants(participantStr);

                        aWorkUnit.weaknessStr = wsSource2?.Cells[row, 12]?.Value?.ToString().Trim();
                        aWorkUnit.strengthStr = wsSource2?.Cells[row, 13]?.Value?.ToString().Trim();
                        aWorkUnit.recommendationStr = wsSource2?.Cells[row, 14]?.Value?.ToString().Trim();

                        aWorkUnit.Char = wsSource2?.Cells[row, 15]?.Value?.ToString().Trim();

                        aPractice.PrjSup_Elements.Add(aWorkUnit);
                        aWorkUnit = new PrjSup_Element();
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

                        aWorkUnit.oE_Elements.Add(aOEelement);
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