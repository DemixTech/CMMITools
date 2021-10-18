using Microsoft.Office.Interop.Excel;
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
    [Serializable]
    public class TargetToolkitObject : TargetFileObject
    {

        const int cToolkitSearchUntilEmptyColumnNotIIGOV = 2;
        const int cToolkitEndTestNotIIGOV = 15; // test for end of file for brute find

        const int cToolkitSearchUntilEmptyColumnIIGOV = 3;
        const int cToolkitEndTestIIGOV = 3; // test for end of file for brute find

        const int cToolkitHeadingStartRow = 1;
        const int cToolkitMaxRows = 10000;

        public override bool LoadPersistantXMLdata()
        {
            try
            {
                // base.LoadPersistant(); override the base function, to load all information from here for this object and its parent
                if (File.Exists(_directoryFileNameXML))
                {
                    // If the directory and file name exists, laod the data
                    var xs = new XmlSerializer(typeof(TargetToolkitObject));
                    using (FileStream xmlLoad = File.Open(_directoryFileNameXML, FileMode.Open))
                    {
                        var pData = (TargetToolkitObject)xs.Deserialize(xmlLoad);
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
            if (o is TargetToolkitObject tko)
            {
                if (!Directory.Exists(Path.GetDirectoryName(_directoryFileNameXML)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(_directoryFileNameXML));
                }
                var xs = new XmlSerializer(typeof(TargetToolkitObject));
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
            Workbook mainWorkbook;
            if ((mainWorkbook = Helper.CheckIfOpenAndOpenXlsx(_directoryFileName)) == null)
            {
                resultMessage = "File not found, has it been moved or deleted?";
                return false;
            }
            string basePath = Path.GetDirectoryName(_directoryFileName);

            string statusStr = "Toolkit master:";
            lblStatus.Text = statusStr;

            int LastUsedRow = 1;
            foreach (Worksheet wksToolkitMaster in mainWorkbook.Worksheets)
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
                    Range mainRange = wksToolkitMaster.Range["A1", "Z" + LastUsedRow];

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





    }
}
