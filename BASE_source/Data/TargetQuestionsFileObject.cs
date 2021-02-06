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
        
        public List<PracticeArea> CMMIModel2 = new List<PracticeArea>();

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
                if ((questionWorkbook = Helper.CheckIfOpenAndOpen(_directoryFileName)) == null)
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
                return true;

            }


        }
    }
}
