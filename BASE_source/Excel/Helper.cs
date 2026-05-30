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
        private const int cThresholdForEndOfDataInSheet = 7;

        public static int FindEndOfWorksheetBrute(Worksheet aWks, int columToSearch, int firstRow = 1, int lastRow = 50000, int endOfSheetThresholdTest = cThresholdForEndOfDataInSheet)
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
                    if (thresHold >= endOfSheetThresholdTest) return contentFoundAtRow + firstRow - 1;
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
            // Skip ROT enumeration to avoid COM deadlocks
            // Just try to get or create an Excel instance
            if (!File.Exists(pathA))
            {
                return null;
            }

            try
            {
                // Try to get existing Excel instance or create new one
                Microsoft.Office.Interop.Excel.Application excelApp;
                try
                {
                    excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch (COMException)
                {
                    // No Excel instance found, create new one
                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                }

                // Try to find the workbook if already open
                string targetFileName = Path.GetFileName(pathA);
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (string.Equals(wb.Name, targetFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        wb.Application.Visible = true;
                        return wb;
                    }
                }

                // Not open, so open it
                Workbook bWkb = excelApp.Workbooks.Open(pathA);
                bWkb.Application.Visible = true;
                return bWkb;
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"COM Exception in CheckIfOpenAndOpenXlsx: {ex.Message}");
                return null;
            }
        }


        public static Presentation CheckIfOpenAndOpenPptx(string pathA)
        {
            // Skip ROT enumeration to avoid COM deadlocks
            if (!File.Exists(pathA))
            {
                return null;
            }

            try
            {
                // Try to get existing PowerPoint instance or create new one
                Microsoft.Office.Interop.PowerPoint.Application pptxApp;
                try
                {
                    pptxApp = (Microsoft.Office.Interop.PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                }
                catch (COMException)
                {
                    // No PowerPoint instance found, create new one
                    pptxApp = new Microsoft.Office.Interop.PowerPoint.Application();
                }

                // Try to find the presentation if already open
                string targetFileName = Path.GetFileName(pathA);
                foreach (Presentation pres in pptxApp.Presentations)
                {
                    if (string.Equals(pres.Name, targetFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        pres.Application.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue;
                        return pres;
                    }
                }

                // Not open, so open it
                Presentation bPptx = pptxApp.Presentations.Open(pathA);
                bPptx.Application.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue;
                return bPptx;
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"COM Exception in CheckIfOpenAndOpenPptx: {ex.Message}");
                return null;
            }
        }

        // *********************************************************************************
        // List all open excel workbooks
        // *********************************************************************************

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        // https://stackoverflow.com/questions/35366658/retrieve-excel-application-from-process-id/35368963
        public static List<Workbook> ExcelGetRunningOjbects()
        {
            List<Workbook> listOfWorkbooks = new List<Workbook>();

            try
            {
                var excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                foreach (Workbook workbook in excelApp.Workbooks)
                {
                    listOfWorkbooks.Add(workbook);
                }
            }
            catch (COMException)
            {
                // No active Excel instance or COM unavailable
            }

            return listOfWorkbooks;
        }

        public static List<Presentation> PptxGetRunningOjbects()
        {
            List<Presentation> listOfPresentations = new List<Presentation>();

            try
            {
                var pptxApp = (Microsoft.Office.Interop.PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                foreach (Presentation presentation in pptxApp.Presentations)
                {
                    listOfPresentations.Add(presentation);
                }
            }
            catch (COMException)
            {
                // No active PowerPoint instance or COM unavailable
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
