using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace BASE.Data
{
    [Serializable]
    public class TargetCASFileObject : TargetFileObject
    {
        private const int cProjectHeadingStartRow2 = 2; // tab:Projects start row
        private const int cSupportHeadingStartRow2 = 2; // tab:Support start row
        private const int cStaffHeadingStartRow2 = 2; // tab:Staff start row
        private const int cSchedule2HeadingStartRow2 = 1; // tab:Schedule2 heading row

        private const int CProcessListStartColumn = 29; // Start column for processing processes defined
        private const int CProcessListHeaderRow = 2; // The header column to look for the name of the processess

        public Organisation Organisation { get; set; }
        public OrganizationalUnit OrganizationalUnit { get; set; }

        public List<WorkUnit> WorkUnitList2 = new List<WorkUnit>(); // Contain all the work unit detail

        // In a staff list, names can be duplicated accross multiople projects, but WorkID-Name will be unqiue
        public List<Staff> StaffList2 = new List<Staff>(); // Contian all the participant
        public List<OUProcess> OUProcessesList2 = new List<OUProcess>(); // Contain the processess and all their projects. Should only be initialised after WorkUnitList is established
        public List<Schedule2> Schedule2List2 = new List<Schedule2>(); // Contains the schedule 2 information


        //public Dictionary<string, int> DictionaryOUPracticeAreas = new Dictionary<string, int>(); // This list is defined in Tab:BASE and defines the offset (key) to a practice area
        public List<String> DictionaryOUPracticeAreas = new List<string>();
        private const int CBASE_OUPracticeAreaCol = 1;
        private const int CBASE_OUPracticeAreaStratRow = 3;

        public List<String> DictionarySupportHeadings = new List<string>();
        private const int CBASE_SupportHeadingsCol = 2;
        private const int CBASE_SupportHeadingsRowStart = 3;

        public List<String> DictionaryProjectHeadings = new List<string>();
        private const int CBASE_ProjectHeadingsCol = 3;
        private const int CBASE_ProjectHeadingsRowStart = 3;



        //public override void InitialiseObject(string directoryFileNameXML,
        //    System.Windows.Forms.Label labelDirectoryNameXML,
        //    System.Windows.Forms.Label labelFileNameXML,
        //    System.Windows.Forms.Label labelDirectoryName,
        //    System.Windows.Forms.Label labelFileName)
        //{
        //    base.InitialiseObject(directoryFileNameXML, labelDirectoryNameXML, labelFileNameXML, labelDirectoryName, labelFileName);
        //}


        public override void SavePersistant(object o)
        {
            if (o is TargetCASFileObject tc)
            {
                if (!Directory.Exists(Path.GetDirectoryName(_directoryFileNameXML)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(_directoryFileNameXML)); ;
                }

                var xs = new XmlSerializer(typeof(TargetCASFileObject));
                using (FileStream stream = File.Create(_directoryFileNameXML))
                {
                    xs.Serialize(stream, tc);
                }

            } else
            {
                throw new NotImplementedException("Object missmatched");

            }
        }

        //public void SavePersistant(TargetCASFileObject theTargetObject)
        //{
        //}

        public bool LoadCASFile()
        {
            Workbook aWorkbook;

            Cursor.Current = Cursors.WaitCursor;
            if ((aWorkbook = Helper.CheckIfOpenAndOpen(_directoryFileName)) == null)
            {
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
                MessageBox.Show("File not found, has it been moved or deleted?");
                return false;
            }

            // *** file was opened.

            // Step 1: Clear the list to start afresh
            WorkUnitList2.Clear();
            StaffList2.Clear();

            // Step 2: Open the spreadhseet and process it
            Worksheet projectWks = aWorkbook.Sheets["Project&Support"];
            int row = cProjectHeadingStartRow2 + 1;
            string sValue2 = projectWks.Cells[row, 1].Value2;
            while (!string.IsNullOrEmpty(sValue2))
            {
                // Process the list
                WorkUnit aNewWorkUnitItem;
                char firstChar = sValue2.ToUpper()[0];

                switch (firstChar)
                {
                    case 'P':
                        aNewWorkUnitItem = new WorkUnit()
                        {
                            WorkType = EWorkType.project,
                        };
                        aNewWorkUnitItem.AddWorkType(EWorkType.project, projectWks, row, cProjectHeadingStartRow2);
                        break;

                    case 'S':
                        aNewWorkUnitItem = new WorkUnit()
                        {
                            WorkType = EWorkType.support,
                        };
                        aNewWorkUnitItem.AddWorkType(EWorkType.support, projectWks, row, cProjectHeadingStartRow2);
                        break;

                    default:
                        aNewWorkUnitItem = new WorkUnit()
                        {
                            WorkType = EWorkType.support,
                        };
                        break;
                }
                WorkUnitList2.Add(aNewWorkUnitItem);



                row++;
                sValue2 = projectWks.Cells[row, 1].Value2;
            }

            // Step 3: Create the process list
            OUProcessesList2.Clear();

            // Start at col 29 (AC) and search to the right until you find END
            int columnX = CProcessListStartColumn; // 29;
            int headerRow = CProcessListHeaderRow; // 2 Row where the processes are defined (below this row is the marking for the projects)
            string cellProcess = projectWks.Cells[headerRow, columnX].Value;
            int lastRowToProcess = Helper.FindEndOfWorksheet(projectWks, 1, 3, 100);
            while (cellProcess != "END")
            {
                // Load the process name
                OUProcess aProcess = new OUProcess();
                aProcess.Name = cellProcess;

                // Find asssociated projects
                for (int rowX = 3; rowX <= lastRowToProcess; rowX++)
                {
                    string cellMarkedX = projectWks.Cells[rowX, columnX]?.Value;
                    if (cellMarkedX?.ToLower() == "x")
                    { // Marked x, proxcess it
                        string workIdStr = projectWks.Cells[rowX, 1]?.Value.ToLower();
                        // The workId must be valid, cannot be null or empty
                        if (string.IsNullOrEmpty(workIdStr))
                        {
                            MessageBox.Show($"WorkID at {rowX} cannot be null or empty");
                        }
                        else
                        {
                            // Use the workIdStr to find the WorkUnit and attach it to the process
                            WorkUnit aWorkunit = WorkUnitList2.Find(x => x.ID.ToLower() == workIdStr);
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
                OUProcessesList2.Add(aProcess);

                // Test for endless loop
                if (columnX++ > 100)
                {
                    MessageBox.Show("END not found. See if end is listed in Row 2 of Projects tab!");
                    break;
                }
                else
                {
                    cellProcess = projectWks.Cells[headerRow, columnX].Value;
                }
            } // Process until end is found

            // Step 4: Open the participant spreadhseet and process it
            Worksheet participantWks = aWorkbook.Sheets["Staff"];
            row = cStaffHeadingStartRow2 + 1;
            string sValue5 = participantWks.Cells[row, 1].Value2;
            while (!string.IsNullOrEmpty(sValue5))
            {
                // Process the list
                Staff aNewParticipant = new Staff();
                aNewParticipant.StaffAdd(participantWks, row, cStaffHeadingStartRow2);
                if (aNewParticipant.WorkID != null) StaffList2.Add(aNewParticipant);

                row++;
                sValue5 = participantWks.Cells[row, 1].Value2;
            }

            // Step 5: *** Load OU information
            Worksheet casP1Wks = aWorkbook.Sheets["P1PA-SS"];
            this.Organisation = new Organisation();
            Organisation.Name = casP1Wks.Cells[16, 2]?.Value?.ToString() ?? "No name";
            Organisation.AddressLine1 = casP1Wks.Cells[18, 2]?.Value?.ToString() ?? "Address line 1";
            Organisation.AddressLine2 = casP1Wks.Cells[19, 2]?.Value?.ToString() ?? "Address line 2";
            Organisation.City = casP1Wks.Cells[20, 2]?.Value?.ToString() ?? "City";
            Organisation.State = casP1Wks.Cells[21, 2]?.Value?.ToString() ?? "State";
            Organisation.ZipCode = casP1Wks.Cells[22, 2]?.Value?.ToString() ?? "ZipCode";
            Organisation.Country = casP1Wks.Cells[23, 2]?.Value?.ToString() ?? "Country";

            this.OrganizationalUnit = new OrganizationalUnit();
            OrganizationalUnit.Name = casP1Wks.Cells[31, 2]?.Value?.ToString() ?? "OU Name";
            string MaturityLevelStr = casP1Wks.Cells[75, 2]?.Value?.ToString() ?? "Maturity Level 1";
            //System.Text.RegularExpressions.Regex
            string MaturityLevelNumberStr = Regex.Match(MaturityLevelStr, @"\d+").Value;
            OrganizationalUnit.MaturityLevel = int.Parse(MaturityLevelNumberStr);

            // Step 4: Load Scheduel 2
            Schedule2List2.Clear();

            Worksheet schedule2Wks = aWorkbook.Sheets["Schedule2"];
            int NumberOfRows = Helper.FindEndOfWorksheet(schedule2Wks, 1, cSchedule2HeadingStartRow2, 200);
            for (int rowS = cSchedule2HeadingStartRow2 + 1; rowS <= NumberOfRows; rowS++)
            {
                // Process the Schedule 2 list
                Schedule2 aNewSchedule2Record = new Schedule2();
                aNewSchedule2Record.Schedule2Add(schedule2Wks, rowS);

                if (aNewSchedule2Record.WorkID != null) Schedule2List2.Add(aNewSchedule2Record);
            }


            // Step 5 : Load Base lookups
            Worksheet baseLookupWks = aWorkbook.Sheets["BASE"];
            int NumberOfRowsInCol1 = Helper.FindEndOfWorksheet(baseLookupWks, CBASE_OUPracticeAreaCol, CBASE_OUPracticeAreaStratRow, 50);
            DictionaryOUPracticeAreas.Clear();
            for (int aRow = CBASE_OUPracticeAreaStratRow; aRow <= NumberOfRowsInCol1; aRow++)
            {
                string PAstr = baseLookupWks.Cells[aRow, CBASE_OUPracticeAreaCol].Value.ToString().ToUpper().Trim();
                DictionaryOUPracticeAreas.Add(PAstr);// .Add(PAstr, aRow - CBASE_OUPracticeAreaStratRow);
            }


            // Step 6 : Load support lookups
            NumberOfRowsInCol1 = Helper.FindEndOfWorksheet(baseLookupWks, CBASE_SupportHeadingsCol, CBASE_SupportHeadingsRowStart, 50);
            DictionarySupportHeadings.Clear();
            for (int aRow = CBASE_SupportHeadingsRowStart; aRow <= NumberOfRowsInCol1; aRow++)
            {
                string headingStr = baseLookupWks.Cells[aRow, CBASE_SupportHeadingsCol].Value.ToString().Trim();
                DictionarySupportHeadings.Add(headingStr);// .Add(PAstr, aRow - CBASE_OUPracticeAreaStratRow);
            }

            // Step 7 : Load project lookups
            NumberOfRowsInCol1 = Helper.FindEndOfWorksheet(baseLookupWks, CBASE_ProjectHeadingsCol, CBASE_ProjectHeadingsRowStart, 50);
            DictionaryProjectHeadings.Clear();
            for (int aRow = CBASE_ProjectHeadingsRowStart; aRow <= NumberOfRowsInCol1; aRow++)
            {
                string headingStr = baseLookupWks.Cells[aRow, CBASE_ProjectHeadingsCol].Value.ToString().Trim();
                DictionaryProjectHeadings.Add(headingStr);// .Add(PAstr, aRow - CBASE_OUPracticeAreaStratRow);
            }

            //public List<String> DictionarySupportHeadings = new List<string>();
            //private const int CBASE_SupportHeadingsCol = 2;
            //private const int CBASE_SupportHeadingsRowStart = 3;

            //public List<String> DictionaryProjectHeadings = new List<string>();
            //private const int CBASE_ProjectHeadingsCol = 2;
            //private const int CBASE_ProjectHeadingsRowStart = 3;



            // *** Done

            Cursor.Current = Cursors.Default;
            MessageBox.Show($"CAS plan information loaded\nProjects={WorkUnitList2.Where(x => x.WorkType == EWorkType.project).Count()}\n" +
                $"Support={WorkUnitList2.Where(x => x.WorkType == EWorkType.support).Count()}\n" +
                $"Processes={OUProcessesList2.Count()}\n" +
                $"Staff entries={StaffList2.Count()}\nSchedule2 entries={Schedule2List2.Count()}");

            return true;


        }


        public bool ReloadSchedule2()
        {
            Workbook aWorkbook;
            Cursor.Current = Cursors.WaitCursor;
            if ((aWorkbook = Helper.CheckIfOpenAndOpen(_directoryFileName)) == null)
            {
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
                MessageBox.Show("File not found, has it been moved or deleted?");
                return false;
            }

            // Step 4: Load Scheduel 2
            Schedule2List2.Clear();

            Worksheet schedule2Wks = aWorkbook.Sheets["Schedule2"];
            int NumberOfRows = Helper.FindEndOfWorksheet(schedule2Wks, 1, cSchedule2HeadingStartRow2, 200);
            for (int rowS = cSchedule2HeadingStartRow2 + 1; rowS <= NumberOfRows; rowS++)
            {
                // Process the Schedule 2 list
                Schedule2 aNewSchedule2Record = new Schedule2();
                aNewSchedule2Record.Schedule2Add(schedule2Wks, rowS);

                if (aNewSchedule2Record.WorkID != null) Schedule2List2.Add(aNewSchedule2Record);
            }


            Cursor.Current = Cursors.Default;
            MessageBox.Show($"Staff entries={StaffList2.Count()}\nSchedule2 entries={Schedule2List2.Count()}");

            return true;
        }

        public bool Generate_OUParticipants(bool insertRole)
        {
            Workbook aWorkbook;
            Cursor.Current = Cursors.WaitCursor;
            if ((aWorkbook = Helper.CheckIfOpenAndOpen(_directoryFileName)) == null)
            {
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
                MessageBox.Show("File not found, has it been moved or deleted?");
                return false;
            }

            // *** Reload schedule 2
            Schedule2List2.Clear();

            Worksheet schedule2Wks = aWorkbook.Sheets["Schedule2"];
            int NumberOfRows = Helper.FindEndOfWorksheet(schedule2Wks, 1, cSchedule2HeadingStartRow2, 200);
            for (int rowS = cSchedule2HeadingStartRow2 + 1; rowS <= NumberOfRows; rowS++)
            {
                // Process the Schedule 2 list
                Schedule2 aNewSchedule2Record = new Schedule2();
                aNewSchedule2Record.Schedule2Add(schedule2Wks, rowS);

                if (aNewSchedule2Record.WorkID != null) Schedule2List2.Add(aNewSchedule2Record);
            }

            // *** Populate OUParticipants
            Worksheet ouPartWks = aWorkbook.Sheets["C_OUParticipants"];
            ouPartWks.Cells.Clear();
            ouPartWks.Cells[1, 1].Value = "Participant's Full Name (required)";
            ouPartWks.Cells[1, 2].Value = "Participant's Role (required)";
            ouPartWks.Cells[1, 3].Value = "Function (required)";
            int rowX = 2;


            //https://www.codegrepper.com/code-examples/csharp/c%23+distinct+comparer+multiple+properties
            List<Schedule2> DistincSchedule2List = Schedule2List2
                .GroupBy(p => new { p.ParticipantName, p.Role })
                .Select(g => g.First())
                .ToList();

            //var DistincSchedule2List = Schedule2List2.Distinct(new CompareNameAndRole()).ToList();
            foreach (var aSchedule2Record in DistincSchedule2List)
            {
                ouPartWks.Cells[rowX, 1].Value = aSchedule2Record.ParticipantName; // "Participant's Full Name (required)";
                ouPartWks.Cells[rowX, 2].Value = aSchedule2Record.Role; // "Participant's Role (required)";
                ouPartWks.Cells[rowX, 3].Value = OrganizationalUnit.Name; // "Function (required)";
                rowX++;
            }


            // Build OUProjects
            int ColOffset;
            if (insertRole == true)
            {
                ColOffset = 3;
            }
            else
            {
                ColOffset = 2;
            }
            Worksheet ouProjects = aWorkbook.Sheets["C_OUProjects"];
            ouProjects.Cells.Clear();
            ouProjects.Cells[1, 1].Value = "Participant Name";
            if (insertRole) ouProjects.Cells[1, 2].Value = "ROLE-DELETE";

            int colOUX = 0;
            foreach (var aWU in WorkUnitList2)
            {
                ouProjects.Cells[1, ColOffset + colOUX].Value = aWU.Name;
                colOUX++;
            }

            int rowOUPX = 2;
            foreach (var aSchedule2Record in DistincSchedule2List)
            {
                ouProjects.Cells[rowOUPX, 1].Value = aSchedule2Record.ParticipantName; // "Participant's Full Name (required)";
                if (insertRole) ouProjects.Cells[rowOUPX, 2].Value = aSchedule2Record.Role; // Participant role
                var theProjectS = WorkUnitList2.Where(x => x.ID == aSchedule2Record.WorkID);
                foreach (var aProject in theProjectS)
                {
                    int projectOffset = WorkUnitList2.FindIndex(x => x.ID == aProject.ID);
                    ouProjects.Cells[rowOUPX, ColOffset + projectOffset].Value = "X";
                }
                rowOUPX++;
            }
            if (insertRole) ouProjects.Range["B1", $"B{rowOUPX - 1}"].Interior.ColorIndex = 6; // Yellow

            // Build OUPAs
            Worksheet ouPAs = aWorkbook.Sheets["C_OUPracticeAreas"];
            ouPAs.Cells.Clear();
            ouPAs.Cells[1, 1].Value = "Participant Name";
            if (insertRole) ouPAs.Cells[1, 2].Value = "ROLE-DELETE";

            for (int i = 0; i < DictionaryOUPracticeAreas.Count; i++)
            {
                ouPAs.Cells[1, ColOffset + i].Value = DictionaryOUPracticeAreas[i]; // aPA.Value].Value = aPA.Key;
            }

            int rowOUPAX = 2;
            foreach (var aSchedule2Record in DistincSchedule2List)
            {
                ouPAs.Cells[rowOUPAX, 1].Value = aSchedule2Record.ParticipantName; // "Participant's Full Name (required)";
                if (insertRole) ouPAs.Cells[rowOUPAX, 2].Value = aSchedule2Record.Role; //Participant role

                // Note that normally the staffRoleList should only contain 1 entry, but in case the staff role is repeated for different projects and different PAs
                var staffRoleList = StaffList2.Where(x => x.Name == aSchedule2Record.ParticipantName && x.Role == aSchedule2Record.Role).ToList();
                foreach (var aStaffEntry in staffRoleList)
                {
                    foreach (var aPA in aStaffEntry.PAlist)
                    {
                        // Make X's next to the corresponding PA
                        int i = DictionaryOUPracticeAreas.FindIndex(x => x == aPA.PAcode.ToString());
                        if (i < 0)
                        {
                            if (insertRole) ouPAs.Cells[rowOUPAX, 2].Value = ouPAs.Cells[rowOUPAX, 2].Value + $"[{aPA.PAcode.ToString()}]"; // Could not find this PA in the list
                            else MessageBox.Show($"Could not find PA {aPA.PAcode.ToString()}");
                        }
                        else
                        {
                            ouPAs.Cells[rowOUPAX, ColOffset + i].Value = "X";
                        }
                    }
                }
                rowOUPAX++;
            }
            if (insertRole) ouPAs.Range["B1", $"B{rowOUPAX - 1}"].Interior.ColorIndex = 6; // Yellow

            Cursor.Current = Cursors.Default;
            MessageBox.Show($"Staff entries={StaffList2.Count()}\nSchedule2 entries={Schedule2List2.Count()}");

            return true;
        }
        public bool Generate_SupportAndProjectCASSheets()
        {
            Workbook aWorkbook;
            Cursor.Current = Cursors.WaitCursor;
            if ((aWorkbook = Helper.CheckIfOpenAndOpen(_directoryFileName)) == null)
            {
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
                MessageBox.Show("File not found, has it been moved or deleted?");
                return false;
            }

            // *** Do support sheet
            Worksheet casSupportWks = aWorkbook.Sheets["C_Support"];
            casSupportWks.Cells.Clear();
            // int NumberOfRows = Helper.FindEndOfWorksheet(casSupportWks, 1, cSchedule2HeadingStartRow2, 200);
            for (int col = 1; col < DictionarySupportHeadings.Count; col++)
            {
                casSupportWks.Cells[1, col].Value = DictionarySupportHeadings[col - 1];
            }

            var supportFunctionsList = WorkUnitList2.Where(x => x.ID.Substring(0, 1).ToUpper() == "S").ToList();
            int sfRow = 2;
            foreach (var aSupportFunction in supportFunctionsList)
            {
                casSupportWks.Cells[sfRow, 1].Value = aSupportFunction.Name;
                casSupportWks.Cells[sfRow, 2].Value = "Support Function";
                casSupportWks.Cells[sfRow, 3].Value = StaffList2.Where(x => x.WorkID == aSupportFunction.ID).Count(); // FTEs
                casSupportWks.Cells[sfRow, 4].Value = aSupportFunction.Description;
                casSupportWks.Cells[sfRow, 5].Value = "No"; // Not sensitive
                casSupportWks.Cells[sfRow, 6].Value = "";
                casSupportWks.Cells[sfRow, 7].Value = "";
                casSupportWks.Cells[sfRow, 8].Value = OrganizationalUnit.Name;
                casSupportWks.Cells[sfRow, 9].Value = "No"; // Not using suppliers
                var Manager = StaffList2.FirstOrDefault(x => x.WorkID == aSupportFunction.ID && x.Role.ToLower().Contains("manager")); // FTEs
                casSupportWks.Cells[sfRow, 10].Value = Manager?.Name ?? "Unknown";
                casSupportWks.Cells[sfRow, 11].Value = Organisation.AddressLine1;
                casSupportWks.Cells[sfRow, 12].Value = Organisation.AddressLine2;
                casSupportWks.Cells[sfRow, 13].Value = Organisation.City;
                casSupportWks.Cells[sfRow, 14].Value = Organisation.State;
                casSupportWks.Cells[sfRow, 15].Value = Organisation.ZipCode;
                casSupportWks.Cells[sfRow, 16].Value = Organisation.Country;
                casSupportWks.Cells[sfRow, 17].Value = "";

                sfRow++;
            }

            // *** Do project sheet
            Worksheet casProjectWks = aWorkbook.Sheets["C_Projects"];
            casProjectWks.Cells.Clear();
            // int NumberOfRows = Helper.FindEndOfWorksheet(casSupportWks, 1, cSchedule2HeadingStartRow2, 200);
            for (int col = 1; col < DictionaryProjectHeadings.Count; col++)
            {
                casProjectWks.Cells[1, col].Value = DictionaryProjectHeadings[col - 1];
            }

            var projectList = WorkUnitList2.Where(x => x.ID.Substring(0, 1).ToUpper() == "P").ToList();
            int prjRow = 2;
            foreach (var aProject in projectList)
            {
                casProjectWks.Cells[prjRow, 1].Value = aProject.Name;
                casProjectWks.Cells[prjRow, 2].Value = "";
                casProjectWks.Cells[prjRow, 3].Value = StaffList2.Where(x => x.WorkID == aProject.ID).Count(); // FTEs
                casProjectWks.Cells[prjRow, 4].Value = aProject.Description;
                casProjectWks.Cells[prjRow, 5].Value = "No"; // Not sensitive
                casProjectWks.Cells[prjRow, 6].Value = "";
                casProjectWks.Cells[prjRow, 7].Value = "";
                casProjectWks.Cells[prjRow, 8].Value = OrganizationalUnit.Name;
                casProjectWks.Cells[prjRow, 9].Value = "Same as business ojbective.";
                casProjectWks.Cells[prjRow, 10].Value = aProject.Lifecycle;
                casProjectWks.Cells[prjRow, 11].Value = aProject.StartDate;
                casProjectWks.Cells[prjRow, 12].Value = aProject.EndDate;

                casProjectWks.Cells[prjRow, 13].Value = "No"; // Not using suppliers
                var Manager = StaffList2.FirstOrDefault(x => x.WorkID == aProject.ID && x.Role.ToLower().Contains("manager")); // FTEs
                casProjectWks.Cells[prjRow, 14].Value = Manager?.Name ?? "Unknown";

                casProjectWks.Cells[prjRow, 15].Value = Organisation.AddressLine1;
                casProjectWks.Cells[prjRow, 16].Value = Organisation.AddressLine2;
                casProjectWks.Cells[prjRow, 17].Value = Organisation.City;
                casProjectWks.Cells[prjRow, 18].Value = Organisation.State;
                casProjectWks.Cells[prjRow, 19].Value = Organisation.ZipCode;
                casProjectWks.Cells[prjRow, 20].Value = Organisation.Country;
                casProjectWks.Cells[prjRow, 21].Value = "";

                prjRow++;
            }



            Cursor.Current = Cursors.Default;
            MessageBox.Show($"Support entries={supportFunctionsList.Count()}\nProject entries={projectList.Count()}");

            return true;
        }



        public bool CreateSchedule1()
        {
            Workbook aWorkbook;

            Cursor.Current = Cursors.WaitCursor;
            if ((aWorkbook = Helper.CheckIfOpenAndOpen(_directoryFileName)) == null)
            {
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
                MessageBox.Show("File not found, has it been moved or deleted?");
                return false;
            }

            // Step 4: Show schedule
            Worksheet schedule = aWorkbook.Sheets["Schedule"];
            schedule.Cells.Clear();
            schedule.Cells[1, 1].Value = "WorkID";
            schedule.Cells[1, 2].Value = "Work name";
            schedule.Cells[1, 3].Value = "PA";
            schedule.Cells[1, 4].Value = "Participant Name";
            schedule.Cells[1, 5].Value = "Role";
            schedule.Cells[1, 6].Value = "WordID2";
            schedule.Cells[1, 7].Value = "Included";
            // https://www.e-iceblue.com/Tutorials/Spire.XLS/Spire.XLS-Program-Guide/Set-Excel-Background-Color-with-C-VB.NET.html
            // schedule.Range["A1:A6"].Style.Color = Color.BlueViolet;

            // *** For each project selected PA, find the participants that acted in that role
            int outRow = 2;
            List<Schedule1Entry> includedList = new List<Schedule1Entry>();
            List<Schedule1Entry> excludedList = new List<Schedule1Entry>();

            foreach (var workUnit in WorkUnitList2)
            {
                var listOfSampledPAs = workUnit.PAlist.Where(x => x.SampleType == ESampleType.added || x.SampleType == ESampleType.sampled);
                foreach (var aSampledPA in listOfSampledPAs)
                {
                    // This is all the sampled PAs
                    var workUnitStaffList = StaffList2.Where(x => x.WorkID == workUnit.ID);
                    foreach (var workUnitParticipant in workUnitStaffList)
                    {
                        // Find the practice areas that match
                        var participantForSampledWorkUnitPA = workUnitParticipant.PAlist.Where(x => x.PAcode == aSampledPA.PAcode);
                        bool found;
                        if (participantForSampledWorkUnitPA.Count() == 0)
                        { // nothing found here
                            found = false; // nothing found
                        }
                        else
                        {
                            // proces this list
                            found = true;
                            schedule.Cells[outRow, 1].Value = workUnit.ID.ToString();
                            schedule.Cells[outRow, 2].Value = workUnit.Name.ToString();
                            schedule.Cells[outRow, 3].Value = aSampledPA.PAcode.ToString();
                            schedule.Cells[outRow, 4].Value = workUnitParticipant.Name;
                            schedule.Cells[outRow, 5].Value = workUnitParticipant.Role;
                            schedule.Cells[outRow, 6].Value = workUnitParticipant.WorkID;

                            // check if the workUnit.ID and PAcode is not already in the list, if it is, then it is a unecessary duplicate. If duplicate, skip include
                            Schedule1Entry aSchedule1Entry = new Schedule1Entry()
                            {
                                ID = workUnit.ID.ToString(),
                                Name = workUnit.Name.ToString(),
                                PAcode = aSampledPA.PAcode.ToString(),
                                ParticipantName = workUnitParticipant.Name,
                                ParticipantRole = workUnitParticipant.Role,
                                WorkIDcheck = workUnitParticipant.WorkID,
                                //  include = false, set below to make reading more clear
                            };

                            int inListIndex = includedList.FindIndex(x => x.ID == aSchedule1Entry.ID && x.PAcode == aSchedule1Entry.PAcode); ;
                            if (inListIndex < 0)
                            { // not in list, insert in includedList
                                aSchedule1Entry.include = true;
                                includedList.Insert(~inListIndex, aSchedule1Entry);
                                schedule.Cells[outRow, 7].Value = "x";
                            }
                            else
                            { // already included, put in excludedList
                                aSchedule1Entry.include = false;
                                excludedList.Add(aSchedule1Entry);
                                schedule.Cells[outRow, 7].Value = "";
                            }
                            outRow++;


                        }
                    }
                }
            }

            // *** Find distinct participants
            //Worksheet responsibilities = aWorkbook.Sheets["Responsibilities"];
            //responsibilities.Cells.Clear();
            //int respRow = 2;
            //var distinctParticipants = StaffList2.Select(x => x.Name)
            //    .Distinct()
            //    .OrderBy(q => q)
            //    .ToList();

            //List<string> projectNameList = new List<string>();
            //List<string> projectWorkIDList = new List<string>();
            //List<EPAcode> practiceAreaList = new List<EPAcode>();
            //foreach (var distincParticipant in distinctParticipants)
            //{
            //    // first clear the names and practice list
            //    projectNameList.Clear();
            //    projectWorkIDList.Clear();
            //    practiceAreaList.Clear();

            //    // *** List all the Projects
            //    var participantSubset = StaffList2.Where(x => x.Name == distincParticipant);
            //    foreach (var aParticipant in participantSubset)
            //    {
            //        // Add a project if it does not exists
            //        var x = projectNameList.BinarySearch(aParticipant.WorkName);
            //        if (x < 0)
            //        { // Not in list, add it
            //            projectNameList.Insert(~x, aParticipant.WorkName);
            //            projectWorkIDList.Insert(~x, aParticipant.WorkID);
            //        }
            //        else
            //        {
            //            // in list, ignore it
            //        }
            //        // Add the PAs if it does not exist
            //        foreach (var aPa in aParticipant.PAlist)
            //        {
            //            var pai = practiceAreaList.BinarySearch(aPa.PAcode);
            //            if (pai < 0)
            //            {
            //                // Not in list, add it
            //                practiceAreaList.Insert(~pai, aPa.PAcode);
            //            }
            //            else
            //            {
            //                // In list, ignore it
            //            }
            //        }
            //    }

            //    // for this participant, output the practicenames and practice areas to the spreadhseet
            //    responsibilities.Cells[respRow++, 1].Value = distincParticipant;
            //    responsibilities.Cells[respRow++, 2].Value = "Project/Work";
            //    for (int i = 0; i < projectNameList.Count(); i++)
            //    //foreach (var aprojName in projectNameList)
            //    {
            //        //responsibilities.Cells[respRow++, 3].Value = aprojName.ToString();
            //        // projectWorkIDList
            //        responsibilities.Cells[respRow++, 3].Value =
            //            projectWorkIDList[i] + " " + projectNameList[i];
            //        //aprojName.ToString();
            //    }
            //    responsibilities.Cells[respRow++, 2].Value = "Practice Area";
            //    foreach (var aPA in practiceAreaList)
            //    {
            //        responsibilities.Cells[respRow++, 3].Value = aPA.ToString();
            //    }

            //}

            MessageBox.Show("Draft Schedule completed");
            return true;
        }

        public override bool LoadPersistantXMLdata()
        {
            try
            {
                // base.LoadPersistant(); override the base function, to load all information from here for this object and its parent
                if (File.Exists(_directoryFileNameXML))
                {
                    // If the directory and file name exists, laod the data
                    var xs = new XmlSerializer(typeof(TargetCASFileObject));
                    using (FileStream xmlLoad = File.Open(_directoryFileNameXML, FileMode.Open))
                    {
                        var pData = (TargetCASFileObject)xs.Deserialize(xmlLoad);
                        this.DirectoryFileName = pData._directoryFileName;

                        // *** TargetCASFielObject fields and properties
                        //     this.Org = pData.Org;
                        //     this.OU = pData.OU;

                        this.WorkUnitList2 = pData.WorkUnitList2;
                        this.StaffList2 = pData.StaffList2;
                        this.OUProcessesList2 = pData.OUProcessesList2;
                        this.Schedule2List2 = pData.Schedule2List2;
                        this.Organisation = pData.Organisation;
                        this.OrganizationalUnit = pData.OrganizationalUnit;
                        this.DictionaryOUPracticeAreas = pData.DictionaryOUPracticeAreas;
                        this.DictionarySupportHeadings = pData.DictionarySupportHeadings;
                        this.DictionaryProjectHeadings = pData.DictionaryProjectHeadings;

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


        //public override bool LoadFileExcelFileData(string fileNameKeyWord)
        //{
        //    throw new NotImplementedException();
        //}

        //public override bool LoadFileData(string fileNameKeyWord)
        //{
        //    throw new NotImplementedException();
        //}
    }

    //public class CompareNameAndRole : IEqualityComparer<Schedule2>
    //{
    //    public bool Equals(Schedule2 x, Schedule2 y)
    //    {


    //        if (x.ParticipantName == y.ParticipantName)
    //        {
    //            if (x.Role == y.Role) return true;
    //            else return false;
    //        }
    //        else return false;

    //        //throw new NotImplementedException();
    //    }

    //    public int GetHashCode(Schedule2 obj)
    //    {
    //        return obj.GetHashCode();
    //        //  throw new NotImplementedException();
    //    }
    //}
}
