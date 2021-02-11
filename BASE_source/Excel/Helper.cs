using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using BASE.Data;

namespace BASE
{
    static public class Helper
    {
        private const int cMAXRowsInAQuestionWorksheet = 2000; // Maximum number of expected rows in a questions worksheet
        private const int cPAwksStartRow = 9; // the start row of the practice area worksheets to prcoess
        private const int cPA_wks_AcronymColumn = 1; // the Acronum column and also the column used to determine the number of rows of the worksheet
        private const int cPA_wks_IntentColumn = 2; // The column for the intent statment
        private const int cPA_wks_ValueColumn = 2; // The column for the value statment



        /// <summary>
        /// Find the end of the worksheet based on a columToSearch, a firstRow and lastRow
        /// </summary>
        /// <param name="aWks">The worksheet to search</param>
        /// <param name="columToSearch">The column to search (it must have filled to the last row)</param>
        /// <param name="firstRow">The first row to search</param>
        /// <param name="lastRow">The last row to search</param>
        /// <returns></returns>
        public static int FindEndOfWorksheet(Worksheet aWks, int columToSearch, int firstRow = 1, int lastRow = 50000)
        {
            if (lastRow == 1) return lastRow;
            if (Math.Abs(lastRow - firstRow) <= 1) return firstRow;
            var centerX = ((lastRow - firstRow) / 2) + firstRow;
            var cellStr = aWks.Cells[centerX, columToSearch].Value;
            if (string.IsNullOrEmpty(cellStr))
            { // middel is empty, look left
                return FindEndOfWorksheet(aWks, columToSearch, firstRow, centerX);
            }
            else
            {
                // look right
                return FindEndOfWorksheet(aWks, columToSearch, centerX, lastRow);
            }

        }
        private const int cThresholdForEndOfDataInSheet = 10;

        public static int FindEndOfWorksheetBrute(Worksheet aWks, int columToSearch, int firstRow = 1, int lastRow = 50000)
        {
            //double aVal = Char.GetNumericValue('A');
            //Char aCol = Convert.ToChar(aVal - 1 + columToSearch);
            //Range startCell = aWks.Range[ [firstRow, GetExcelColumnName(columToSearch) + ];

            //Range mainRangeIIG = wksMain.Range["G" + 2, "R" + NumberOfRowsIIG];

            //Range endCell = aWks.Range[lastRow, columToSearch];
            Range aRangeToSearch = aWks.Range[GetExcelColumnName(columToSearch) + firstRow, GetExcelColumnName(columToSearch) + lastRow];

            int thresHold = 0;
            int contentFoundAtRow = 0;

            for (int currentRow = 1; currentRow <= lastRow - firstRow + 1; currentRow++)
            {
                if (aRangeToSearch[currentRow, 1].Value != null)
                {
                    thresHold = 0;
                    contentFoundAtRow = currentRow;
                }
                else
                {
                    thresHold++;
                    if (thresHold >= cThresholdForEndOfDataInSheet) return contentFoundAtRow + firstRow - 1;
                }

            }
            return contentFoundAtRow + firstRow - 1;
        }

        // https://stackoverflow.com/questions/181596/how-to-convert-a-column-number-e-g-127-into-an-excel-column-e-g-aa
        static public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }


        /// <summary>
        /// Check if open, if open assign it, else if it exists open it, else return null
        /// </summary>
        /// <param name="pathA"></param>
        /// <returns></returns>

        public static Workbook CheckIfOpenAndOpenXlsx(string pathA)
        {

            List<Workbook> listOfWorkbooks;
            listOfWorkbooks = ExcelGetRunningOjbects(); // Get all the open workbooks
            foreach (Workbook aWkb in listOfWorkbooks)
            {
                string pathFileName = Path.Combine(aWkb.Path, aWkb.Name);
                if (string.Compare(pathA, pathFileName.Trim(), ignoreCase: true) == 0)
                { // string matches workbook exists, open it
                    aWkb.Application.Visible = true;
                    return aWkb;
                }
            }
            // Workbook does not exists, open or create it
            if (File.Exists(pathA))
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook bWkb = excelApp.Workbooks.Open(pathA);
                bWkb.Application.Visible = true;
                return bWkb;
            }
            return null;

        }


        public static Presentation CheckIfOpenAndOpenPptx(string pathA)
        {

            List<Presentation> listOfPresentations;
            listOfPresentations = PptxGetRunningOjbects(); // Get all the open workbooks
            foreach (Presentation aPptx in listOfPresentations)
            {
                string pathFileName = Path.Combine(aPptx.Path, aPptx.Name);
                if (string.Compare(pathA, pathFileName.Trim(), ignoreCase: true) == 0)
                { // string matches workbook exists, open it
                    aPptx.Application.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue;
                    return aPptx;
                }
            }
            // Workbook does not exists, open or create it
            if (File.Exists(pathA))
            {
                Microsoft.Office.Interop.PowerPoint.Application pptxApp = new Microsoft.Office.Interop.PowerPoint.Application();
                Presentation bPptx = pptxApp.Presentations.Open(pathA);
                bPptx.Application.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue;
                return bPptx;
            }
            return null;

        }

        // *********************************************************************************
        // List all open excel workbooks
        // *********************************************************************************

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        // https://stackoverflow.com/questions/35366658/retrieve-excel-application-from-process-id/35368963
        public static List<Workbook> ExcelGetRunningOjbects()
        {
            IRunningObjectTable lRunningObjectTable = null;
            IEnumMoniker lMonikerList = null;
            List<Workbook> listOfWorkbooks = new List<Workbook>();

            try
            {
                // Query Running Object Table 
                if (GetRunningObjectTable(0, out lRunningObjectTable) != 0 || lRunningObjectTable == null)
                {
                    return listOfWorkbooks;
                }

                // List Monikers
                lRunningObjectTable.EnumRunning(out lMonikerList);

                // Start Enumeration
                lMonikerList.Reset();

                // Array used for enumerating Monikers
                IMoniker[] lMonikerContainer = new IMoniker[1];

                IntPtr lPointerFetchedMonikers = IntPtr.Zero;

                // foreach Moniker
                while (lMonikerList.Next(1, lMonikerContainer, lPointerFetchedMonikers) == 0)
                {
                    object lComObject;
                    lRunningObjectTable.GetObject(lMonikerContainer[0], out lComObject);

                    // Check the object is an Excel workbook
                    if (lComObject is Microsoft.Office.Interop.Excel.Workbook)
                    {
                        Microsoft.Office.Interop.Excel.Workbook lExcelWorkbook = (Microsoft.Office.Interop.Excel.Workbook)lComObject;
                        // Show the Window Handle 
                        // MessageBox.Show("Found Excel Application with Window Handle " + lExcelWorkbook.Application.Hwnd);
                        listOfWorkbooks.Add(lExcelWorkbook);

                        //MessageBox.Show($"Workbook name {lExcelWorkbook.Name}");
                    }
                }
            }
            finally
            {
                // Release ressources
                if (lRunningObjectTable != null) Marshal.ReleaseComObject(lRunningObjectTable);
                if (lMonikerList != null) Marshal.ReleaseComObject(lMonikerList);
            }
            return listOfWorkbooks;
        }

        public static List<Presentation> PptxGetRunningOjbects()
        {
            IRunningObjectTable lRunningObjectTable = null;
            IEnumMoniker lMonikerList = null;
            List<Presentation> listOfPresentations = new List<Presentation>();

            try
            {
                // Query Running Object Table 
                if (GetRunningObjectTable(0, out lRunningObjectTable) != 0 || lRunningObjectTable == null)
                {
                    return listOfPresentations;
                }

                // List Monikers
                lRunningObjectTable.EnumRunning(out lMonikerList);

                // Start Enumeration
                lMonikerList.Reset();

                // Array used for enumerating Monikers
                IMoniker[] lMonikerContainer = new IMoniker[1];

                IntPtr lPointerFetchedMonikers = IntPtr.Zero;

                // foreach Moniker
                while (lMonikerList.Next(1, lMonikerContainer, lPointerFetchedMonikers) == 0)
                {
                    object lComObject;
                    lRunningObjectTable.GetObject(lMonikerContainer[0], out lComObject);

                    // Check the object is an Excel workbook
                    if (lComObject is Microsoft.Office.Interop.PowerPoint.Presentation)
                    {
                        Microsoft.Office.Interop.PowerPoint.Presentation lPptxPresentation = (Microsoft.Office.Interop.PowerPoint.Presentation)lComObject;
                        // Show the Window Handle 
                        // MessageBox.Show("Found Excel Application with Window Handle " + lExcelWorkbook.Application.Hwnd);
                        listOfPresentations.Add(lPptxPresentation);

                        //MessageBox.Show($"Workbook name {lExcelWorkbook.Name}");
                    }
                }
            }
            finally
            {
                // Release ressources
                if (lRunningObjectTable != null) Marshal.ReleaseComObject(lRunningObjectTable);
                if (lMonikerList != null) Marshal.ReleaseComObject(lMonikerList);
            }
            return listOfPresentations;
        }

        public static PracticeArea ProcessPracticeArea(Workbook questionWkb, string worksheetName)
        {
            PracticeArea aNewPracticeArea = new PracticeArea();

            // Worksheet practiceAreaWks = questionWkb.Worksheets[worksheetName];
            foreach (Worksheet paWks in questionWkb.Worksheets)
            {
                if (paWks.Name.ToLower() == worksheetName.ToLower())
                {
                    // Process this worksheet
                    aNewPracticeArea.PAcode = (EPAcode)Enum.Parse(typeof(EPAcode), worksheetName);
                    // Find the end of the worksheet
                    int workSheetRows = Helper.FindEndOfWorksheet(paWks, cPA_wks_AcronymColumn, cPAwksStartRow, cMAXRowsInAQuestionWorksheet);
                    // Read the name
                    aNewPracticeArea.Name = paWks.Cells[1, 1].Value;
                    aNewPracticeArea.NameChinese = paWks.Cells[2, 1].Value;

                    // Read the intent and value statements
                    aNewPracticeArea.Intent = paWks.Cells[3, cPA_wks_IntentColumn].Value;
                    aNewPracticeArea.IntentChinese = paWks.Cells[4, 2].Value;
                    aNewPracticeArea.Value = paWks.Cells[5, 2].Value;
                    aNewPracticeArea.ValueChinese = paWks.Cells[6, cPA_wks_ValueColumn].Value;

                    // Find the end of the sheet and process each row
                    for (int row = cPAwksStartRow; row <= workSheetRows; row++)
                    {
                        string cellStr = paWks.Cells[row, cPA_wks_AcronymColumn].Value;
                        if (cellStr.ToLower().Contains(worksheetName.ToLower()))
                        {
                            // contains the acronym, process it
                            Practice aPractice = ProcessPractice(paWks, row, cellStr);
                            if (aPractice != null) aNewPracticeArea.Practices.Add(aPractice);
                        }
                    }
                    return aNewPracticeArea;
                }
            }

            return null;
        }

        public static Practice ProcessPractice(Worksheet paWks, int row, string firstColStr)
        {
            // Process the row in paWks
            Practice aPractice = new Practice();
            aPractice.Acronym = paWks.Name;

            // Find the level and the practice number
            Regex rxNumber = new Regex(@"[\.][\d]+");
            Regex rxLevel = new Regex(@"[\d]+[\.]");
            Regex rxValueOnly = new Regex(@"\d+");

            // Test firstCoLStr for the number
            Match match = rxValueOnly.Match(rxNumber.Match(firstColStr).Value.ToString());
            aPractice.Number = int.Parse(match.Value.ToString());

            // Test firstCoLStr for the level
            match = rxValueOnly.Match(rxLevel.Match(firstColStr).Value.ToString());
            aPractice.Level = int.Parse(match.Value.ToString());

            string[] english;
            string[] chinese;

            // Process the practice statement
            string practiceStatement = paWks.Cells[row, 2].Value;
            ExtractEnglishAndChinese(practiceStatement, out english, out chinese);
            aPractice.Statement = english[0];
            aPractice.StatementChinese = chinese[0];

            // Extract the example activity
            string exampleActivityStr = paWks.Cells[row, 3]?.Value ?? null;
            ExtractEnglishAndChinese(exampleActivityStr, out english, out chinese);
            for (int i = 0; i < english?.Length; i++)
            {
                ExampleActivity aActivity = new ExampleActivity()
                {
                    Activity = english[i],
                    ActivityChinese = chinese[i],
                };
                aPractice.ExampleActivities.Add(aActivity);
            }

            // Extract the example work products
            string exampleWorkProductsStr = paWks.Cells[row, 4]?.Value ?? null;
            ExtractEnglishAndChinese(exampleWorkProductsStr, out english, out chinese);
            for (int i = 0; i < english?.Length; i++)
            {
                ExampleWorkProduct aWorkProduct = new ExampleWorkProduct()
                {
                    Description = english[i],
                    DescriptionChinese = chinese[i],
                };
                aPractice.ExampleWorkProducts.Add(aWorkProduct);
            }

            // Extract the questions
            string questions = paWks.Cells[row, 5]?.Value ?? null;
            ExtractEnglishAndChinese(questions, out english, out chinese);
            for (int i = 0; i < english?.Length; i++)
            {
                Question aQuestion = new Question()
                {
                    Sentence = english[i],
                    SentenceChinese = chinese[i],
                };
                aPractice.Questions.Add(aQuestion);
            }


            return aPractice;


        }

        private static void ExtractEnglishAndChinese(string incommingString, out string[] english, out string[] chinese)
        {
            if (incommingString == null)
            {
                english = null;
                chinese = null;
                return;
            }
            List<string> eStr = new List<string>();
            List<string> cStr = new List<string>();
            string[] stringArray = incommingString.Split('-');
            foreach (string stringElement in stringArray)
            {
                string[] englChinese = stringElement.Split('\n');
                if (!string.IsNullOrEmpty(englChinese[0]))
                {
                    eStr.Add(englChinese[0].Trim());
                    if (englChinese.Count() > 1 && !string.IsNullOrEmpty(englChinese[1])) cStr.Add(englChinese[1].Trim());
                    else cStr.Add("");
                }
            }

            english = eStr.ToArray();
            chinese = cStr.ToArray();

        }

        public static Worksheet OpenOrElseCreateWks(Workbook aWkb, string wksNameToOpen)
        {
            // Test if the workbook exists, if it does, return it
            foreach (Worksheet aWks in aWkb.Worksheets)
            {
                if (aWks.Name == wksNameToOpen) return aWks;
            }
            // The worksheet was not found, create one
            Worksheet aNewWks = aWkb.Worksheets.Add();
            aNewWks.Name = wksNameToOpen;
            return aNewWks;
        }

        public static void ExtractPracticeAreaInformation(Practice aPractice,
             out string statementStr, out string workProductStr,
                                out string activityStr, out string questionStr)
        {
            statementStr = aPractice.Statement + "\n" + aPractice.StatementChinese;
            workProductStr = "";
            bool newWorkProduct = true;
            foreach (var aWkp in aPractice.ExampleWorkProducts)
            {
                if (newWorkProduct)
                {
                    workProductStr = "- " + aWkp.Description + "\n" + aWkp.DescriptionChinese;
                    newWorkProduct = false;
                }
                else
                {
                    workProductStr = "\n- " + aWkp.Description + "\n" + aWkp.DescriptionChinese;

                }
            }
            activityStr = "";
            bool newActivity = true;
            foreach (var aActivity in aPractice.ExampleActivities)
            {
                if (newActivity)
                {
                    activityStr = "- " + aActivity.Activity + "\n" + aActivity.ActivityChinese;
                    newActivity = false;
                }
                else
                {
                    activityStr = "\n- " + aActivity.Activity + "\n" + aActivity.ActivityChinese;
                }
            }

            questionStr = "";
            bool newQuestion = true;
            foreach (var aQuestion in aPractice.Questions)
            {
                if (newQuestion)
                {
                    questionStr = "- " + aQuestion.Sentence + "\n" + aQuestion.SentenceChinese;
                    newQuestion = false;
                }
                else
                {
                    questionStr = "\n- " + aQuestion.Sentence + "\n" + aQuestion.SentenceChinese;
                }
            }

        }


        public static Worksheet AssignOrCreateWorksheet(Workbook aWkb, string wksName, string InsertAfterWksStr)
        {
            Worksheet insertWks = FindWorksheet(aWkb, InsertAfterWksStr);
            if (insertWks == null)
            {
                insertWks = aWkb.Worksheets.Add();
                insertWks.Name = InsertAfterWksStr;
            }


            Worksheet newWks = FindWorksheet(aWkb, wksName);
            if (newWks == null)
            {
                newWks = aWkb.Worksheets.Add(insertWks);
                newWks.Name = wksName;

            }

            return newWks;
        }

        public static Worksheet FindWorksheet(Workbook aWkb, string wksName)
        {
            foreach (Worksheet aWks in aWkb.Worksheets)
            {
                if (aWks.Name.ToUpper() == wksName.ToUpper())
                {
                    return aWks;
                }
            }
            return null;
        }

    }
}
