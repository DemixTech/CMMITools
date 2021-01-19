using BASE.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace BASE
{
    public class PersistentData
    {
        private const string CPersistantDataFileName_Generic = "OEpersistent.xml";
        private const string CPersistantDataFileName_Questions = "OEpersistentQuestions.xml";
        private const string CPersistantDataFileName_WorkUnits = "OEpersistentWorkUnits.xml";
        private const string CPersistantDataFileName_ProcessList = "OEpersistentProcessList.xml";
        private const string CPersistantDataFileName_StaffList = "OEpersistentStaffList.xml";
        private const string CPersistantDataFileName_Schedule2List = "OEpersistentSchedule2List.xml";

        public string LastAppraisalDirectory { get; set; }
        public string CASPlanName { get; set; }
        public string QuestionPathFile { get; set; }
        static public string PersistantPathFile_Generic { get; set; }
        static public string PersistantPathFile_Questions { get; set; }
        static public string PersistantPathFile_WorkUnits { get; set; }
        static public string PersistantPathFile_ProcessLists { get; set; }
        static public string PersistantPathFile_StaffLists { get; set; }
        static public string PersistantPathFile_Schedule2Lists { get; set; }


        public string DemixToolPathFile { get; set; } // Demix Tool
        public string DemixTool_ToImport_PathFile { get; set; } // Demix tool to import

        public string AppToolMainPathFile { get; set; } // OEdbMainPathFile => 5MB MDD-Toolkit-Appraisal Tool.xlsm 
        public string AppToolSourcePathFile { get; set; } // OEdbSourcePathFile to import => 5MB MDD-Toolkit-Appraisal Tool.xlsm 

        public string OEdatabasePathFile { get; set; } // Path file to the OE database

        public string FromText { get; set; } // The from text for file renaming
        public string ToText { get; set; }  // The to text for file renaming
        // **** Private data
        private string tempDirectory;

        public PersistentData()
        {
            tempDirectory = Path.GetTempPath();
            PersistantPathFile_Generic = Path.Combine(tempDirectory, CPersistantDataFileName_Generic);
            LoadPersistentData();
            
            PersistantPathFile_Questions = Path.Combine(tempDirectory, CPersistantDataFileName_Questions);
            PersistantPathFile_WorkUnits = Path.Combine(tempDirectory, CPersistantDataFileName_WorkUnits);
            PersistantPathFile_ProcessLists = Path.Combine(tempDirectory, CPersistantDataFileName_ProcessList);
            PersistantPathFile_StaffLists = Path.Combine(tempDirectory, CPersistantDataFileName_StaffList);
            PersistantPathFile_Schedule2Lists = Path.Combine(tempDirectory, CPersistantDataFileName_Schedule2List);
            

        }
        public void LoadPersistentData()
        {
            //string fullPathFileName = Path.Combine(tempDirectory, CPersistantDataFileName);
            try
            {

                if (File.Exists(PersistantPathFile_Generic))
                {
                    var xs = new XmlSerializer(typeof(PersistentData));
                    using (FileStream xmlLoad = File.Open(PersistantPathFile_Generic, FileMode.Open))
                    {
                        var pData = (PersistentData)xs.Deserialize(xmlLoad);
                        this.CASPlanName = pData.CASPlanName;
                        this.QuestionPathFile = pData.QuestionPathFile;
                        this.LastAppraisalDirectory = pData.LastAppraisalDirectory;
                        this.AppToolSourcePathFile = pData.AppToolSourcePathFile;
                        this.AppToolMainPathFile = pData.AppToolMainPathFile;
                        this.OEdatabasePathFile = pData.OEdatabasePathFile;
                        
                        this.DemixToolPathFile = pData.DemixToolPathFile;
                        this.DemixTool_ToImport_PathFile = pData.DemixTool_ToImport_PathFile;

                        this.FromText = pData.FromText;
                        this.ToText = pData.ToText;
                     }
                }
                else
                {
                    Initialise();
                }
            }
            catch (Exception ex)
            {
                Initialise();
            }

        }

       // public List<PracticeArea> CMMIModel = new List<PracticeArea>();

        public static void LoadPersistentData_Questions(ref List<PracticeArea> aCMMIModel)
        {
            //string fullPathFileName = Path.Combine(tempDirectory, CPersistantDataFileName);
            try
            {
                if (File.Exists( PersistantPathFile_Questions))
                {
                    var xs = new XmlSerializer(typeof(List<PracticeArea>));
                    using (FileStream xmlLoad = File.Open(PersistantPathFile_Questions, FileMode.Open))
                    {
                        var pData = (List<PracticeArea>)xs.Deserialize(xmlLoad);
                        aCMMIModel = pData;
                        return;
                    }
                }
                else
                {
                   // Initialise();
                }
            }
            catch (Exception ex)
            {
             //   Initialise();
            }
            return;
        }
        public static void SavePersistentData_Questions(List<PracticeArea> cmmiModel)
        {

           // XmlDocument doc = new XmlDocument();
            var xs = new XmlSerializer(typeof(List<PracticeArea>));
            // Create a file to write to
            using (FileStream stream = File.Create(PersistantPathFile_Questions))
            {
                xs.Serialize(stream, cmmiModel);
            }
        }

        public static void LoadPersistentData_WorkUnitList(ref List<WorkUnit> aWorkUnitList)
        {
            //string fullPathFileName = Path.Combine(tempDirectory, CPersistantDataFileName);
            try
            {
                if (File.Exists(PersistantPathFile_WorkUnits))
                {
                    var xs = new XmlSerializer(typeof(List<WorkUnit>));
                    using (FileStream xmlLoad = File.Open(PersistantPathFile_WorkUnits, FileMode.Open))
                    {
                        var pData = (List<WorkUnit>)xs.Deserialize(xmlLoad);
                        aWorkUnitList = pData;
                        return;
                    }
                }
                else
                {
                    // Initialise();
                }
            }
            catch (Exception ex)
            {
                //   Initialise();
            }
            return;
        }
        public static void SavePersistentData_WorkUnits(List<WorkUnit> aWorkUnitList)
        {

            // XmlDocument doc = new XmlDocument();
            var xs = new XmlSerializer(typeof(List<WorkUnit>));
            // Create a file to write to
            using (FileStream stream = File.Create(PersistantPathFile_WorkUnits))
            {
                xs.Serialize(stream, aWorkUnitList);
            }
        }

        public static void LoadPersistentData_ProcessList(ref List<OUProcess> aProcessList)
        {
            //string fullPathFileName = Path.Combine(tempDirectory, CPersistantDataFileName);
            try
            {
                if (File.Exists(PersistantPathFile_ProcessLists))
                {
                    var xs = new XmlSerializer(typeof(List<OUProcess>));
                    using (FileStream xmlLoad = File.Open(PersistantPathFile_ProcessLists, FileMode.Open))
                    {
                        var pData = (List<OUProcess>)xs.Deserialize(xmlLoad);
                        aProcessList = pData;
                        return;
                    }
                }
                else
                {
                    // Initialise();
                }
            }
            catch (Exception ex)
            {
                //   Initialise();
            }
            return;
        }
        public static void SavePersistentData_ProcessLists(List<OUProcess> aProcessList)
        {

            // XmlDocument doc = new XmlDocument();
            var xs = new XmlSerializer(typeof(List<OUProcess>));
            // Create a file to write to
            using (FileStream stream = File.Create(PersistantPathFile_ProcessLists))
            {
                xs.Serialize(stream, aProcessList);
            }
        }


        public static void LoadPersistentData_StaffList(ref List<Staff> staffList)
        {
            //string fullPathFileName = Path.Combine(tempDirectory, CPersistantDataFileName);
            try
            {
                if (File.Exists(PersistantPathFile_StaffLists))
                {
                    var xs = new XmlSerializer(typeof(List<Staff>));
                    using (FileStream xmlLoad = File.Open(PersistantPathFile_StaffLists, FileMode.Open))
                    {
                        var pData = (List<Staff>)xs.Deserialize(xmlLoad);
                        staffList = pData;
                        return;
                    }
                }
                else
                {
                    // Initialise();
                }
            }
            catch (Exception ex)
            {
                //   Initialise();
            }
            return;
        }
        public static void SavePersistentData_StaffList(List<Staff> staffList)
        {

            // XmlDocument doc = new XmlDocument();
            var xs = new XmlSerializer(typeof(List<Staff>));
            // Create a file to write to
            using (FileStream stream = File.Create(PersistantPathFile_StaffLists))
            {
                xs.Serialize(stream, staffList);
            }
        }

        public static void LoadPersistentData_Schedule2List(ref List<Schedule2> schedule2List)
        {
            //string fullPathFileName = Path.Combine(tempDirectory, CPersistantDataFileName);
            try
            {
                if (File.Exists(PersistantPathFile_Schedule2Lists))
                {
                    var xs = new XmlSerializer(typeof(List<Schedule2>));
                    using (FileStream xmlLoad = File.Open(PersistantPathFile_Schedule2Lists, FileMode.Open))
                    {
                        var pData = (List<Schedule2>)xs.Deserialize(xmlLoad);
                        schedule2List = pData;
                        return;
                    }
                }
                else
                {
                    // Initialise();
                }
            }
            catch (Exception ex)
            {
                //   Initialise();
            }
            return;
        }
        public static void SavePersistentData_Schedule2List(List<Schedule2> schedule2List)
        {

            // XmlDocument doc = new XmlDocument();
            var xs = new XmlSerializer(typeof(List<Schedule2>));
            // Create a file to write to
            using (FileStream stream = File.Create(PersistantPathFile_Schedule2Lists))
            {
                xs.Serialize(stream, schedule2List);
            }
        }

        private void Initialise()
        {
            LastAppraisalDirectory = Environment.CurrentDirectory;
            CASPlanName = "";
            AppToolMainPathFile = Environment.CurrentDirectory;
            AppToolSourcePathFile = Environment.CurrentDirectory;
        }

        public void SavePersistentData(PersistentData pd)
        {
            //persistentData.SavePersistentData();
            // Does the file exists, if not, create it
            //string fullPathFileName = Path.Combine(tempDirectory, CPersistantDataFileName);

            XmlDocument doc = new XmlDocument();

            // if (File.Exists(fullPathFileName))
            //  {
            var xs = new XmlSerializer(typeof(PersistentData));
            // Create a file to write to
            using (FileStream stream = File.Create(PersistantPathFile_Generic))
            {
                xs.Serialize(stream, pd);
            }


        }

    }


}
