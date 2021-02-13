using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace BASE.Data
{
    /// <summary>
    /// TargetFile is an object that contains information for a file, including saving its content to XML and loading it whe it starts up
    /// </summary>
    [Serializable]
    abstract public class TargetFileObject
    {
        public const string CCASinName = "CAS"; // This string must be present in the filename (not the directory) to be valid
        public const string COEdbinName = "OEdbATL"; // This string must be present in the filename (not the directory) to be valid
        public const string COEdbATMinName = "OEdbATM"; // The team leads version of the OEdb
        public const string CQuestionInName = "Question"; // Question is in the name

        // https://docs.microsoft.com/en-us/dotnet/api/system.xml.serialization.xmlattributes.xmlignore?view=net-5.0
        //https://docs.microsoft.com/en-us/dotnet/api/system.xml.serialization.xmlattributes.xmlignore?view=net-5.0#System_Xml_Serialization_XmlAttributes_XmlIgnore

        [XmlIgnoreAttribute]
        protected Label _labelFileName;
        [XmlIgnoreAttribute]
        protected Label _labelDirecotryName;

        public string _directoryFileName = "";
        [XmlIgnoreAttribute]
        public string DirectoryFileName
        {

            get
            {
                return _directoryFileName;
            }
            set
            {
                _directoryFileName = value;
                _labelFileName.Text = Path.GetFileName(_directoryFileName);
                _labelDirecotryName.Text = Path.GetDirectoryName(_directoryFileName);
            }
        }

        [XmlIgnoreAttribute]
        protected Label _labelFileNameXML;
        [XmlIgnoreAttribute]
        protected Label _labelDirectoryNameXML;

        public string _directoryFileNameXML = "";

        [XmlIgnoreAttribute]
        public string DirectoryFileNameXML
        {

            get
            {
                return _directoryFileNameXML;
            }
            set
            {
                _directoryFileNameXML = value;
                _labelFileNameXML.Text = Path.GetFileName(_directoryFileNameXML);
                _labelDirectoryNameXML.Text = Path.GetDirectoryName(_directoryFileNameXML);
            }
        }


        public virtual void InitialiseObject(string directoryFileNameXML,
            Label labelDirectoryNameXML, Label labelFileNameXML,
            Label labelDirectoryName, Label labelFileName)
        {

            _labelDirectoryNameXML = labelDirectoryNameXML;
            _labelFileNameXML = labelFileNameXML;
            DirectoryFileNameXML = directoryFileNameXML;

            _labelDirecotryName = labelDirectoryName;
            _labelFileName = labelFileName;

        }
        // Load information about this TargetFileObject
        abstract public bool LoadPersistantXMLdata();
        abstract public void SavePersistant(object o);

        public virtual void ClearPathFile()
        {
            DirectoryFileName = "";
            SavePersistant(this);
        }


        /// <summary>
        /// Open a dialog box and select a file (not open it, but select it)
        /// </summary>
        /// <returns></returns>

        public bool SelectFileToLoad(string fileNameKeyWord)

        //  virtual public bool LoadFileData(string fileNameKeyWord)
        {
            // Check if the excel process is running

            OpenFileDialog sourceFile2 = new OpenFileDialog();

            // *** _directoryFileName = null or "" then default to working directoyr
            if (String.IsNullOrEmpty(_directoryFileName))
            {
                sourceFile2.InitialDirectory = Directory.GetCurrentDirectory(); // persistentData.LastAppraisalDirectory; //cPath_start;
            }
            else
            {
                sourceFile2.InitialDirectory = Path.GetDirectoryName(_directoryFileName);
            }
            sourceFile2.RestoreDirectory = true;
            sourceFile2.Title = "Select source file";
            //sourceFile2.DefaultExt = "*.xlsx";
            if (sourceFile2.ShowDialog() == DialogResult.OK)
            {
                // *** If the file name does not contain CAS, then we need to abort
                if (string.IsNullOrEmpty(fileNameKeyWord) || Path.GetFileName(sourceFile2.FileName).ToUpper().Contains(fileNameKeyWord.ToUpper()))
                {
                    // Set cursor as hourglass
                    //Cursor.Current = Cursors.WaitCursor;
                    this.DirectoryFileName = sourceFile2.FileName;
                    // this.SavePersistant();
                    return true;
                }
                else
                { // Filename does not contain CCASinName
                    MessageBox.Show($"The filename {Path.GetFileName(sourceFile2.FileName)} does not contain the keyword \"{fileNameKeyWord}\"");
                    return false;
                }

            }
            else
            {
                return false;
            }
        }

    }
}