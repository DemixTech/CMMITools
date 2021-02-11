using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
//using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
//using DocumentFormat.OpenXml.Office.Drawing;
using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Office.Interop.PowerPoint;
//using System;
//using System.Collections.Generic;
//using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;


namespace BASE.Data
{
    public class TargetPresentationFileObject : TargetFileObject
    {
        public override bool LoadPersistantXMLdata()
        {
            try
            {
                // base.LoadPersistant(); override the base function, to load all information from here for this object and its parent
                if (File.Exists(_directoryFileNameXML))
                {
                    // If the directory and file name exists, laod the data
                    var xs = new XmlSerializer(typeof(TargetPresentationFileObject)); // TargetCASFileObject));
                    using (FileStream xmlLoad = File.Open(_directoryFileNameXML, FileMode.Open))
                    {
                        var pData = (TargetPresentationFileObject)xs.Deserialize(xmlLoad);
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
            if (o is TargetPresentationFileObject tc)
            {
                if (!Directory.Exists(Path.GetDirectoryName(_directoryFileNameXML)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(_directoryFileNameXML)); ;
                }

                var xs = new XmlSerializer(typeof(TargetPresentationFileObject));
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


        public bool UpdateLinks(TargetDataReferenceFileObject BASEDataReferenceObject)
        {
            if (BASEDataReferenceObject == null || string.IsNullOrEmpty(BASEDataReferenceObject._directoryFileName))
            {
                MessageBox.Show($"The Data reference file has not be selected!");
                return false;
            }

            //Presentation aPresentation;

            Cursor.Current = Cursors.WaitCursor;
            //if ((aPresentation = Helper.CheckIfOpenAndOpenPptx(_directoryFileName)) == null)
            //{
            //    // Set cursor as default arrow
            //    Cursor.Current = Cursors.Default;
            //    MessageBox.Show("File not found, has it been moved or deleted?");
            //    return false;
            //}

            //// *** file is open, do something
            //string listOfLInks;
            ////foreach (var aLink in aPresentation.)
            ///

            string listOfOEs = "";
            // Returns all the external hyperlinks in the slides of a presentation.
            //public static IEnumerable<String> GetAllExternalHyperlinksInPresentation(string fileName)
            // {
            // Declare a list of strings.
            List<string> ret = new List<string>();

            // Open the presentation file as read-only.
            // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.presentationdocument?view=openxml-2.8.1

            string fileLinksStr = "";
            using (PresentationDocument document = PresentationDocument.Open(_directoryFileName, false))
            {

                // https://stackoverflow.com/questions/28460333/inserting-an-image-in-footer-using-openxml-in-net-c-sharp
                // Iterate through all the slide parts in the presentation part.
                foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                {
                    //vhttp://www.ericwhite.com/blog/forums/topic/adding-workbook-to-chart-as-an-embeddedpackagepart/
                    List<ExternalRelationship> references = slidePart.ExternalRelationships.ToList();
                    foreach (var aRef in references)
                    {
                        fileLinksStr += $"{aRef.Uri}\n";
                    }

                }
            }

            // Copy the pptx

            // rename the current one to .zip


            // find the contents in the zip and rename it
            // https://www.example-code.com/csharp/zip_update_file.asp
            // https://docs.microsoft.com/en-us/dotnet/api/system.io.compression.zipfile.openread?view=net-5.0
            // https://docs.microsoft.com/en-us/dotnet/api/system.io.compression.brotlistream?view=net-5.0





            // Return the list of strings.
            // return ret;
            //}
            MessageBox.Show($"The links are \n{fileLinksStr}");
            return true;
        }

        // Get the presentation object and pass it to the next CountSlides method.
        private static SlideId GetHiddenSlide(PresentationDocument presentationDocument)
        {
            // Open the presentation as read-only.
            SlideId slideId = null;
            foreach (SlidePart slide in presentationDocument.PresentationPart.SlideParts)
            {
                if (slide.Slide.Show != null && !slide.Slide.Show)
                {
                    string slideRelId = presentationDocument.PresentationPart.GetIdOfPart(slide);
                    // Get the slide ID of the specified slide
                    slideId = presentationDocument.PresentationPart.Presentation.SlideIdList.ChildElements.Where(
                        s => ((SlideId)s).RelationshipId == slideRelId).FirstOrDefault() as SlideId;
                    break;
                }
            }
            // Pass the presentation to the next CountSlide method
            // and return the slide count.
            return slideId;
        }
    }
}
