using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using Label = Microsoft.Office.Interop.Excel.Label;

namespace BASE.Data
{
    [Serializable]
    public class TargetQuestionsFileObject : TargetFileObject
    {
        public string Myname = "Pieter van ZYl";

        public List<PracticeArea> CMMIModel2 = new List<PracticeArea>(); // Defines the full CMMI model with practices, questions, artifacts and activities

        public List<MapRecord> MapRecords = new List<MapRecord>(); // Map records define the layout of characterisation map

        const int CtmpStartRow = 4; // exclude heading at 3
        const int CtmpEndRow = 35;
        const int CtmpStartCol = 3; // exclude Practice nubmer at 2
        const int CtmpEndCol = 21;

        public override bool LoadPersistantXMLdata()
        {
            try
            {
                // base.LoadPersistant(); override the base function, to load all information from here for this object and its parent
                if (File.Exists(_directoryFileNameXML))
                {
                    // If the directory and file name exists, laod the data
                    var xs = new XmlSerializer(typeof(TargetQuestionsFileObject)); // TargetCASFileObject));
                    using (FileStream xmlLoad = File.Open(_directoryFileNameXML, FileMode.Open))
                    {
                        var pData = (TargetQuestionsFileObject)xs.Deserialize(xmlLoad);
                        this.DirectoryFileName = pData._directoryFileName;

                        // *** Load the object elements belwo
                        this.Myname = pData.Myname;
                        this.CMMIModel2 = pData.CMMIModel2;
                        this.MapRecords = pData.MapRecords;
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
            if (o is TargetQuestionsFileObject tc)
            {
                if (!Directory.Exists(Path.GetDirectoryName(_directoryFileNameXML)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(_directoryFileNameXML)); ;
                }

                var xs = new XmlSerializer(typeof(TargetQuestionsFileObject));
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

        public bool LoadTheQuestionAndModelFile(System.Windows.Forms.Label lblStatus2)
        {
            // *** Test if the question file exists
            if (!File.Exists(_directoryFileName))
            {
                MessageBox.Show($"The question file {_directoryFileName}\ndoes not exists!");
                return false;
            }
            else
            {
                Workbook questionWorkbook;
                if ((questionWorkbook = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
                {
                    // *** 
                    MessageBox.Show("File not found, has it been moved or deleted?");
                    return false;
                }

                // Clear the model and start processin the questionWorkbook
                CMMIModel2.Clear();

                //MessageBox.Show($"The question file exists, now processing it ... ");
                string statusStr = "";
                foreach (var worksheetName in Enum.GetValues(typeof(EPAcode))) //.Cast<EPAcode>().ToList())
                {
                    // Open the worksheet and process it
                    PracticeArea aPracticeArea = Helper.ProcessPracticeArea(questionWorkbook, worksheetName.ToString());
                    if (aPracticeArea != null) CMMIModel2.Add(aPracticeArea);

                    // Update status string
                    statusStr += worksheetName + " ";
                    lblStatus2.Text = statusStr;
                }

            }
            BuildMapRecords();
            return true;
        }

        private void BuildMapRecords()
        {
            MapRecords.Clear();

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

                    // *** Build the map from the spreadhseet *** Enhancemenet build the map from the CMMI Model
                    string cellStr = tmpWks.Cells[rowX, colY]?.Value?.ToString() ?? ""; // Note this can be OOS, -, or something else
                    bool OoSValue = false;
                    if (tmpWks.Cells[rowX, colY]?.Value == "OoS") OoSValue = true;
                    string RowColStr = Helper.GetExcelColumnName(colY) + rowX.ToString();

                    if (!string.IsNullOrEmpty(cellStr))
                    {
                        MapRecord aMapRecord = new MapRecord()
                        {
                            PAstr = PAstr,
                            Col = colY,
                            Row = rowX,
                            LevelStr = numberStr,
                            OoS = OoSValue,
                            RowColStr = RowColStr,
                            PALevelStr = PAstr + " " + numberStr,
                        };
                        MapRecords.Add(aMapRecord);
                    }
                }
            }
        }

    }
}
