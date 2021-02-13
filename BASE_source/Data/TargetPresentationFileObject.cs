using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
//using System.IO.Compression;
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

            Cursor.Current = Cursors.WaitCursor;

            List<string> searchStringList = new List<string>();
            List<string> onlyPathFileList = new List<string>();

            // Open the presentation file as read-only.
            // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.presentationdocument?view=openxml-2.8.1

            List<ExternalRelationship> relationShipList = new List<ExternalRelationship>();
            using (PresentationDocument document = PresentationDocument.Open(_directoryFileName, false))
            {
                //foreach (var aPart in document.ExtendedFilePropertiesPart.DataPartReferenceRelationships) // .GetPartsOfType<TitlesOfParts>())
                //{
                //    // aPart.tit

                //}

                // https://stackoverflow.com/questions/28460333/inserting-an-image-in-footer-using-openxml-in-net-c-sharp
                // Iterate through all the slide parts in the presentation part.
                foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                {
                    //vhttp://www.ericwhite.com/blog/forums/topic/adding-workbook-to-chart-as-an-embeddedpackagepart/
                    List<ExternalRelationship> references = slidePart.ExternalRelationships.ToList();
                    foreach (var aRef in references)
                    {
                        relationShipList.Add(aRef);
                        string searchStr = aRef.Uri.ToString();
                        searchStr = searchStr.Substring(8, searchStr.Length - 8);

                        int resultIndex = searchStringList.FindIndex(x => x == searchStr);
                        if (resultIndex < 0) searchStringList.Insert(~resultIndex, searchStr);

                        string pathFileStr = "";
                        int pptxIndx = searchStr.IndexOf(".xlsm");
                        if (pptxIndx >= 0)
                        {
                            // *** replace
                            pathFileStr = searchStr.Substring(0, pptxIndx);
                            pathFileStr = pathFileStr + ".xlsm";

                        }
                        else
                        { // path index not found, look for xlsx
                            pptxIndx = searchStr.IndexOf(".xlsx");
                            if (pptxIndx >= 0)
                            {
                                // *** replace
                                pathFileStr = searchStr.Substring(0, pptxIndx);
                                pathFileStr = pathFileStr + ".xlsx";
                            }
                        }
                        if (pptxIndx >= 0 && !string.IsNullOrEmpty(pathFileStr))
                        {
                            int resultIndex2 = onlyPathFileList.FindIndex(x => x == pathFileStr);
                            if (resultIndex2 < 0) onlyPathFileList.Insert(~resultIndex2, pathFileStr);
                        }

                    }

                }
            }

            // find the contents in the zip and rename it
            // https://www.example-code.com/csharp/zip_update_file.asp
            // https://docs.microsoft.com/en-us/dotnet/api/system.io.compression.zipfile.openread?view=net-5.0
            // https://docs.microsoft.com/en-us/dotnet/api/system.io.compression.brotlistream?view=net-5.0

            // https://docs.microsoft.com/en-us/office/open-xml/how-to-search-and-replace-text-in-a-document-part

            string fileLinksStr = "";
            foreach (string uStr in onlyPathFileList)
            {
                fileLinksStr += $"{uStr}\n";
            }
            //MessageBox.Show($"The links are \n{fileLinksStr}");

            SearchAndReplace(_directoryFileName, onlyPathFileList, BASEDataReferenceObject._directoryFileName);

            return true;
        }


        // To search and replace content in a document part.
        public static void SearchAndReplace(string _directoryFileName2, List<string> searchStringList2, string replacementStr)
        {
            // *** Make a copy of orrignial
            string pathStr = Path.GetDirectoryName(_directoryFileName2);
            string fileNoExt = Path.GetFileNameWithoutExtension(_directoryFileName2);
            string extStr = Path.GetExtension(_directoryFileName2);
            string dateTimeStr = DateTime.Now.ToString("g").Replace("/", "").Replace(":", ""); ;

            string backupFileStr = Path.Combine(pathStr, fileNoExt + dateTimeStr + extStr);
            int fileSuffix = 0;
            while (File.Exists(backupFileStr))
            {
                fileSuffix++;
                backupFileStr = Path.Combine(pathStr, fileNoExt + dateTimeStr + "_" + fileSuffix + extStr);
            }
            File.Copy(_directoryFileName2, backupFileStr);

            // *** Reading a zip file
            string ListOfZips = "All entries\n";

            // https://docs.telerik.com/devtools/document-processing/libraries/radziplibrary/features/update-ziparchive
            try
            {
                using (Stream stream = File.Open(_directoryFileName2, FileMode.Open))
                {
                    using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Update, false, null))
                    {
                        // Display the list of the files in the selected zip file using the ZipArchive.Entries property. 
                        foreach (ZipArchiveEntry zipFile in archive.Entries)
                        {
                            ListOfZips += zipFile.FullName + "\n";

                            if (zipFile.FullName.Length > 17 && (zipFile.FullName.Substring(0, 17) == "ppt/slides/_rels/"))
                            {

                                // *** open the file and update it 
                                Stream entryStream = zipFile.Open();
                                StreamReader reader = new StreamReader(entryStream);
                                string docText = reader.ReadToEnd();

                                // *** find and replace content
                                //string contentReplaced = content.Replace("line", "<replaced line>");
                                //if (string.IsNullOrEmpty(contentReplaced)) { contentReplaced = "My line to insert."; }
                                foreach (string aString in searchStringList2)
                                {
                                    var cleanString = aString.Replace("/", @"\"); // FindAndSubstituteAll(aString, " ", "%20");
                                                                                  //cleanString = FindAndSubstituteAll(cleanString, "/", @"\");
                                    int locationIndex = 0;
                                    do
                                    {
                                        locationIndex = FindAndSubstituteSeachString(ref docText, locationIndex, cleanString, replacementStr);
                                        if (locationIndex > 0)
                                        {
                                            string stopHere = "";
                                        }
                                    } while (locationIndex >= 0 && (locationIndex < docText.Length));

                                }

                                //entryStream.Seek(0, SeekOrigin.End);
                                // *** write content back
                                entryStream.Seek(0, SeekOrigin.Begin);
                                StreamWriter writer = new StreamWriter(entryStream);
                                writer.WriteLine(docText);
                                writer.Flush();
                            }
                        }
                    }

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Message:{ex.Message}");
            }


          //  MessageBox.Show($"List of zips\n {ListOfZips}");
        }



        public static void SearchAndReplaceOld(string _directoryFileName2, List<string> searchStringList2, string replacementStr)
        {
            // *** Make a copy of orrignial
            string pathStr = Path.GetDirectoryName(_directoryFileName2);
            string fileNoExt = Path.GetFileNameWithoutExtension(_directoryFileName2);
            string extStr = Path.GetExtension(_directoryFileName2);
            string dateTimeStr = DateTime.Now.ToString("g").Replace('/', '-');

            string backupFileStr = Path.Combine(pathStr, fileNoExt + dateTimeStr + extStr);
            int fileSuffix = 0;
            while (File.Exists(backupFileStr))
            {
                fileSuffix++;
                backupFileStr = Path.Combine(pathStr, fileNoExt + dateTimeStr + "_" + fileSuffix + extStr);
            }

            File.Copy(_directoryFileName2, backupFileStr);

            // *** Reading a zip file
            string ListOfZips = "";
            // https://docs.microsoft.com/en-us/dotnet/api/system.io.compression.zipfile.open?view=net-5.0
            using (ZipArchive zip = ZipFile.Open(_directoryFileName2, ZipArchiveMode.Update))
            {
                foreach (var zipFile in zip.Entries)
                {
                    ListOfZips += zipFile.FullName + "\n";
                    // open the file and update it 
                    string docText = null;

                    //var aZipFile = zipFile.Open();
                    //var docTex2t = aZipFile.Read();

                    //if (zipFile.FullName == "ppt/slides/_rels/slide4.xml.rels")
                    //{
                    //    string breakHere = "stop";
                    //}
                    using (StreamReader sr = new StreamReader(zipFile.Open()))//zipFile.Open.ExtendedFilePropertiesPart.GetStream(FileMode.Open))) //.PresentationPart.GetStream(FileMode.Open)))
                    {
                        docText = sr.ReadToEnd();
                    }

                    // *** find all in the list and replace it
                    foreach (string aString in searchStringList2)
                    {
                        var cleanString = FindAndSubstituteAll(aString, " ", "%20");
                        cleanString = FindAndSubstituteAll(cleanString, "/", @"\");
                        int locationIndex = 0;
                        do
                        {

                            locationIndex = FindAndSubstituteSeachString(ref docText, locationIndex, cleanString, replacementStr);
                            if (locationIndex > 0)
                            {
                                string stopHere = "";
                            }
                        } while (locationIndex >= 0 && (locationIndex < docText.Length));

                    }

                    using (StreamWriter sw = new StreamWriter(zipFile.Open())) // pptDoc.ExtendedFilePropertiesPart.GetStream(FileMode.Create))) // wordDoc.PresentationPart.GetStream(FileMode.Create))) // .MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }

                }

            }


            MessageBox.Show($"List of zips\n {ListOfZips}");

            #region remove
            ////   using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            using (PresentationDocument pptDoc = PresentationDocument.Open(_directoryFileName2, true))
            {
                //    #region TestThis
                //    // test find and replace

                //    //string testStr1 = "daar was <geen probeelm> met die";
                //    //string replaceStr1 = "<alles ok>";

                //    //int nextLocation = FindAndSubstituteSeachString(ref testStr1, 0, "<geen probeelm>", replaceStr1);
                //    //MessageBox.Show($"The updated string {testStr1}\nNext location={nextLocation}");

                //    //string testStr2 = "<geen probeelm>";
                //    //string replaceStr2 = "222";

                //    //int nextLocation2 = FindAndSubstituteSeachString(ref testStr2, 0, "<geen probeelm>", replaceStr2);
                //    //MessageBox.Show($"The updated string {testStr2}\nNext location={nextLocation2}");

                //    //string testStr3 = "asdasdsa <geen probeelm dfgdgsdf";
                //    //string replaceStr3 = "777";

                //    //int nextLocation3 = FindAndSubstituteSeachString(ref testStr3, 0, "<geen probeelm>", replaceStr3);
                //    //MessageBox.Show($"The updated string {testStr3}\nNext location={nextLocation3}");
                //    //string docText = null;
                //    //using (StreamReader sr = new StreamReader(pptDoc.CoreFilePropertiesPart.GetStream(FileMode.Open))) //.PresentationPart.GetStream(FileMode.Open)))
                //    //{
                //    //    docText = sr.ReadToEnd();
                //    //}

                //    //Regex regexText = new Regex(@"S:/2021-02-20to02-26 (A5) C54321 ShortName/00_Data_Reference.xlsm", RegexOptions.IgnoreCase);
                //    //var matches = regexText.Matches(docText);

                //    //docText = regexText.Replace(docText, $@"S:/2021-02-20to02-26 (A5) R416 D5406 C51828 Maxvision/00_Data_Reference.xlsm");

                //    //using (StreamWriter sw = new StreamWriter(pptDoc.CoreFilePropertiesPart.GetStream(FileMode.Create))) // wordDoc.PresentationPart.GetStream(FileMode.Create))) // .MainDocumentPart.GetStream(FileMode.Create)))
                //    //{
                //    //    sw.Write(docText);
                //    //}
                #endregion

                string docText = null;
                using (StreamReader sr = new StreamReader(pptDoc.ExtendedFilePropertiesPart.GetStream(FileMode.Open))) //.PresentationPart.GetStream(FileMode.Open)))
                {
                    docText = sr.ReadToEnd();
                }

                // *** find all in the list and replace it
                foreach (string aString in searchStringList2)
                {
                    var cleanString = FindAndSubstituteAll(aString, " ", "%20");
                    cleanString = FindAndSubstituteAll(cleanString, "/", @"\");
                    int locationIndex = 0;
                    do
                    {
                        locationIndex = FindAndSubstituteSeachString(ref docText, locationIndex, aString.Replace('/', '\\'), replacementStr);
                        if (locationIndex > 0)
                        {
                            string stopHere = "";
                        }
                    } while (locationIndex >= 0 && (locationIndex < docText.Length));

                }

                using (StreamWriter sw = new StreamWriter(pptDoc.ExtendedFilePropertiesPart.GetStream(FileMode.Create))) // wordDoc.PresentationPart.GetStream(FileMode.Create))) // .MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }


                //foreach (SlidePart slidePart in pptDoc.PresentationPart.SlideParts)
                //// foreach (var aPart in pptDoc.part .Parts)
                //{

                //    // *** update the slidePart
                //    docText = null;
                //    using (StreamReader sr = new StreamReader(slidePart.GetStream(FileMode.Open))) //.PresentationPart.GetStream(FileMode.Open)))
                //    {
                //        docText = sr.ReadToEnd();
                //    }
                //    // *** find all in the list and replace it
                //    foreach (string aString in searchStringList2)
                //    {
                //        int locationIndex = 0;
                //        do
                //        {
                //            locationIndex = FindAndSubstituteSeachString(ref docText, locationIndex, aString.Replace('/', '\\'), replacementStr);
                //            if (locationIndex > 0)
                //            {
                //                string stopHere = "";
                //            }
                //        } while (locationIndex >= 0 && (locationIndex < docText.Length));

                //    }

                //    using (StreamWriter sw = new StreamWriter(slidePart.GetStream(FileMode.Create))) // wordDoc.PresentationPart.GetStream(FileMode.Create))) // .MainDocumentPart.GetStream(FileMode.Create)))
                //    {
                //        sw.Write(docText);
                //    }
            }

            //    //foreach (SlidePart slidePart in pptDoc.PresentationPart.GetPartsOfType<SlidePart>())
            //    foreach (SlidePart slidePart in pptDoc.PresentationPart.GetPartsOfType<SlidePart>())
            //    // foreach (var aPart in pptDoc.part .Parts)
            //    {

            //        // *** update the slidePart
            //        docText = null;
            //        using (StreamReader sr = new StreamReader(slidePart.GetStream(FileMode.Open))) //.PresentationPart.GetStream(FileMode.Open)))
            //        {
            //            docText = sr.ReadToEnd();
            //        }
            //        // *** find all in the list and replace it
            //        foreach (string aString in searchStringList2)
            //        {
            //            int locationIndex = 0;
            //            do
            //            {
            //                locationIndex = FindAndSubstituteSeachString(ref docText, locationIndex, aString.Replace('/', '\\'), replacementStr);
            //                if (locationIndex > 0)
            //                {
            //                    string stopHere = "True";
            //                }
            //            } while (locationIndex >= 0 && (locationIndex < docText.Length));

            //        }

            //        using (StreamWriter sw = new StreamWriter(slidePart.GetStream(FileMode.Create))) // wordDoc.PresentationPart.GetStream(FileMode.Create))) // .MainDocumentPart.GetStream(FileMode.Create)))
            //        {
            //            sw.Write(docText);
            //        }
            //    }





            //}
        }
        private static string FindAndSubstituteAll(string incommingStr, string searchString, string replacementString)
        {
            int startIndex = 0;
            do
            {
                startIndex = FindAndSubstituteSeachString(ref incommingStr, startIndex, searchString, replacementString);
            } while (startIndex >= 0 && (startIndex < incommingStr.Length));
            return incommingStr;
        }
        ///// <summary>
        /// Find from locaton in searchText the startTag and endTag. Then replace everyting between startTag and endTag with replacementText.
        /// Return -1 if nothing found, return the location of the endTag's last character
        /// </summary>
        /// <param name="searchText"></param>
        /// <param name="location"></param>
        /// <param name="startTag"></param>
        /// <param name="endTag"></param>
        /// <param name="replacementText"></param>
        /// <returns></returns>
        private static int FindAndSubstituteSeachString(ref string sourceString, int startIndex2, string searchString, string replacementStr)
        {

            int startIndexLeft = sourceString.IndexOf(searchString, startIndex2);
            if (startIndexLeft < 0) return -1; // nothing found
            string leftString = sourceString.Substring(0, startIndexLeft); //, searchText.Length - startIndexLeft);

            // else something found, now find the end tag
            int endIndexRight = startIndexLeft + searchString.Length;
            string rightString = sourceString.Substring(endIndexRight, sourceString.Length - endIndexRight);
            string returnString = leftString + replacementStr + rightString;
            sourceString = returnString;
            return startIndexLeft + replacementStr.Length; // + endTag.Length; // endIndexRight;
        }


        /// <summary>
        /// Find from locaton in searchText the startTag and endTag. Then replace everyting between startTag and endTag with replacementText.
        /// Return -1 if nothing found, return the location of the endTag's last character
        /// </summary>
        /// <param name="searchText"></param>
        /// <param name="location"></param>
        /// <param name="startTag"></param>
        /// <param name="endTag"></param>
        /// <param name="replacementText"></param>
        /// <returns></returns>
        private static int FindAndSubstituteBetweenTags(ref string searchText, int startIndex2, string startTag, string endTag, string replacementText)
        {
            int startIndexLeft = searchText.IndexOf(startTag, startIndex2);
            if (startIndexLeft < 0) return -1; // nothing found
            string leftString = searchText.Substring(0, startIndexLeft + startTag.Length); //, searchText.Length - startIndexLeft);

            // else something found, now find the end tag
            int endIndexRight = searchText.IndexOf(endTag, startIndexLeft + startTag.Length);
            if (endIndexRight < 0) return -2; // could not find the end tag
                                              // else start tag found and end tag found
            string rightString = searchText.Substring(endIndexRight, searchText.Length - endIndexRight);
            string returnString = leftString + replacementText + rightString;
            searchText = returnString;
            return startIndexLeft + startTag.Length + replacementText.Length + endTag.Length; // endIndexRight;
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
