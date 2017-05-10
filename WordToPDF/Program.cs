using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp;
using System.IO;
//using iTextSharp.text;
//using iTextSharp.text.pdf;
//using Spire.Doc;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Diagnostics;

namespace WordToPDF
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Enter 1 for Single Document Process Or 2 For Directory Process");
            string option = Console.ReadLine();
            if (option == "1")
            {
                Console.WriteLine("Hello Drag and Drop the Word Doc file onto the screen and press enter");
                string filepath = Console.ReadLine().Replace("\"", "");
                string newFilePath = filepath.Remove(filepath.LastIndexOf('.'));
                newFilePath = string.Concat(newFilePath, DateTime.Now.ToString().Replace(":", "_").Replace("/", "") + ".pdf").Replace(" ", "_");
                if (File.Exists(filepath))
                {
                    Word.ApplicationClass application = new Word.ApplicationClass();
                    Word.Document Doc = application.Documents.Open(filepath);

                    string success = SingleDoc(application, Doc, newFilePath);
                    Console.WriteLine(success);
                    Console.WriteLine("Press Enter To Finish");
                    Console.ReadLine();

                    #region OldCode
                    //int fileLen;
                    //FileStream stream = File.Open(filepath, FileMode.Open);
                    //BinaryReader reader = new BinaryReader(stream);
                    //fileLen = reader.ReadInt32();
                    ////long fileLen = stream.Length;
                    //Byte[] Input = new Byte[fileLen];

                    //stream.Read(Input, 0, fileLen);
                    //Document doc = new Document();
                    #endregion

                    //object oMissing = System.Reflection.Missing.Value;

                    //Document Doc = Word.Documents.Open(filepath, ref oMissing, ref oMissing, ref oMissing
                    //    , ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                    //    , ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                    //    , ref oMissing, ref oMissing);

                    //try
                    //{
                    //    ExtractMetadata2XML(Doc);
                    //}
                    //catch (Exception ex)
                    //{
                    //    Doc.Close();
                    //    application.Quit();
                    //    // Handle exception if for some reason the document is not available.
                    //    Debug.WriteLine("procdoc_parse: " + ex.ToString());
                    //}


                    #region OldCode
                    //Spire.Doc.Document sdoc = new Spire.Doc.Document(filepath);
                    //sdoc.SaveToFile(newFilePath, Spire.Doc.FileFormat.PDF);

                    //// create PDF File and create a writer on it
                    //PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(newFilePath, FileMode.Create));
                    //// open the document
                    //doc.Open();
                    //// Add the text file contents
                    //doc.Add(new Paragraph(System.Text.Encoding.Default.GetString(Input)));
                    //// Close the document
                    //doc.Close();
                    //// Get the Posted file Content Length
                    //fileLen = fu.PostedFile.ContentLength;
                    //// Create a byte array with content length
                    //Byte[] Input = new Byte[fileLen];
                    //// Create stream
                    //System.IO.Stream myStream;
                    //// get the stream of uploaded file
                    //myStream = fu.FileContent;
                    //// Read from the stream
                    //myStream.Read(Input, 0, fileLen);
                    // Create a Document
                    //Document doc = new Document();
                    //// create PDF File and create a writer on it
                    //PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream
                    //(string.Concat(Server.MapPath("~/Pdf/PdfSample"),
                    //".pdf"), FileMode.Create));
                    //// open the document
                    //doc.Open();
                    //// Add the text file contents
                    //doc.Add(new Paragraph(System.Text.Encoding.Default.GetString(Input)));
                    //// Close the document
                    //doc.Close();
                    #endregion
                }
            }
            else if (option == "2")
            {
                Console.WriteLine("Hello Drag and Drop the Word Doc Directory onto the screen and press enter");
                string directory = Console.ReadLine().Replace("\"", "");
                if (Directory.Exists(directory))
                {
                    string[] docfiles = Directory.GetFiles(directory, "*.doc*", SearchOption.TopDirectoryOnly);
                    foreach (string filepath in docfiles)
                    {
                        string newFilePath = filepath.Remove(filepath.LastIndexOf('.'));
                        newFilePath = string.Concat(newFilePath, DateTime.Now.ToString().Replace(":", "_").Replace("/", "") + ".pdf").Replace(" ", "_");
                        if (File.Exists(filepath))
                        {
                            Word.ApplicationClass application = new Word.ApplicationClass();
                            Word.Document Doc = application.Documents.Open(filepath);

                            string success = SingleDoc(application, Doc, newFilePath);
                            Console.WriteLine(success);
                        }
                    }
                    Console.WriteLine("Press Enter To Finish");
                    Console.ReadLine();
                }
                else if (!Directory.Exists(directory))
                {
                    Console.WriteLine("Directory Does Not Exist, Press Enter To Exit");
                    Console.ReadLine();
                }
            }
        }

        protected static string SingleDoc(Word.ApplicationClass application, Word.Document Doc, string newFilePath)
        {
            try
            {
                SavePDF sv = new SavePDF();
                bool finish = sv.Save(Doc, newFilePath);
            }
            catch (Exception ex)
            {
                Doc.Close();
                application.Quit();
                Console.WriteLine("Save PDF Failed With: " + ex.ToString());
                Console.WriteLine("Press Enter To Continue/Exit");
                Console.ReadLine();
            }
            Doc.Close();
            application.Quit();
            return "Save PDF Finished Successfully";
        }

        private bool ParseCRITable(Word.Table tb, string doctype, string filename, string filepath, out Dictionary<string, string> reportdata)
        {
            bool retval = false;
            string x = "";
            reportdata = new Dictionary<string, string>();
            reportdata.Add("filename", filename);
            reportdata.Add("path", filepath);
            reportdata.Add("doctype", doctype);

            if (tb.Rows.Count == 7 && tb.Columns.Count == 4)
            {
                // title
                x = tb.Rows[(int)MetadataTableRows.htitle].Range.Text;
                if (!x.Contains("Title")) return (retval);
                reportdata.Add("title", tb.Rows[(int)MetadataTableRows.title].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty));

                // author
                x = tb.Rows[(int)MetadataTableRows.hauthor].Range.Text;
                if (!x.Contains("Author")) return (retval);
                reportdata.Add("authors", tb.Rows[(int)MetadataTableRows.author].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty));

                // final reviewer
                x = tb.Rows[(int)MetadataTableRows.hfinalreviewer].Range.Text;
                if (!x.Contains("Final Reviewer")) return (retval);

                // final reviewer
                x = tb.Cell((int)MetadataTableRows.finalreviewer, (int)MetadataTableCols.finalreviewer).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                reportdata.Add("finalreviewer", x);

                // eln prj number
                x = tb.Cell((int)MetadataTableRows.finalreviewer, (int)MetadataTableCols.elnprj).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", ",");
                reportdata.Add("elnprj", x);

                // report date
                x = tb.Cell((int)MetadataTableRows.finalreviewer, (int)MetadataTableCols.reportdate).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                reportdata.Add("reportdate", x);

                // report number
                x = tb.Cell((int)MetadataTableRows.finalreviewer, (int)MetadataTableCols.reportnumber).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                reportdata.Add("reportnumber", x);

                retval = true;
            }
            else if (tb.Rows.Count == 7 && tb.Columns.Count == 3) // no eln prj or databook number
            {
                // document type
                //string x = tb.Rows[(int)MetadataTableRows.doctype].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                //if (!x.ToLower().Contains("central report index (cri) report")) return (retval);
                //reportdata.Add("doctype", "CRI");

                // title
                x = tb.Rows[(int)MetadataTableRows.htitle].Range.Text;
                if (!x.Contains("Title")) return (retval);
                reportdata.Add("title", tb.Rows[(int)MetadataTableRows.title].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty));

                // author
                x = tb.Rows[(int)MetadataTableRows.hauthor].Range.Text;
                if (!x.Contains("Author")) return (retval);
                reportdata.Add("authors", tb.Rows[(int)MetadataTableRows.author].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty));

                // final reviewer
                x = tb.Rows[(int)MetadataTableRows.hfinalreviewer].Range.Text;
                if (!x.Contains("Final Reviewer")) return (retval);

                // final reviewer
                x = tb.Cell((int)MetadataTableRows.finalreviewer, (int)MetadataTableCols.finalreviewer).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                reportdata.Add("finalreviewer", x);

                // report date
                x = tb.Cell((int)MetadataTableRows.finalreviewer, (int)MetadataTableCols.elnprj).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                reportdata.Add("reportdate", x);

                // report number
                x = tb.Cell((int)MetadataTableRows.finalreviewer, (int)MetadataTableCols.reportdate).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                reportdata.Add("reportnumber", x);

                retval = true;
            }
            else if (tb.Rows.Count == 9 && tb.Columns.Count == 3)
            {
                // document type
                //string x = tb.Rows[(int)MetadataTableRows.doctype].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                //if (!x.ToLower().Contains("central report index (cri) report")) return (retval);
                //reportdata.Add("doctype", "CRI");

                // title
                x = tb.Rows[(int)MetadataTableRows.htitle].Range.Text;
                if (!x.Contains("Title")) return (retval);
                reportdata.Add("title", tb.Rows[(int)MetadataTableRows.title].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty));

                // author
                x = tb.Rows[(int)MetadataTableRows.hauthor].Range.Text;
                if (!x.Contains("Author")) return (retval);
                reportdata.Add("authors", tb.Rows[(int)MetadataTableRows.author].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty));

                // databook numbers and number of pages
                x = tb.Rows[(int)MetadataTableRows.hfinalreviewer].Range.Text;
                if (!x.Contains("Databook Numbers")) return (retval);

                // data book numbers
                x = tb.Cell((int)MetadataTableRows.finalreviewer, (int)MetadataTableCols.elnprj).Range.Text.Trim().Replace("\r", ",").Replace("\a", string.Empty);
                reportdata.Add("elnprj", x);

                // final reviewer
                x = tb.Rows[(int)MetadataTableRows.hfinalreviewer + 2].Range.Text;
                if (!x.Contains("Final Reviewer")) return (retval);

                // final reviewer
                x = tb.Cell((int)MetadataTableRows.finalreviewer + 2, (int)MetadataTableCols.finalreviewer).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                reportdata.Add("finalreviewer", x);

                // report date
                x = tb.Cell((int)MetadataTableRows.finalreviewer + 2, (int)MetadataTableCols.elnprj).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                reportdata.Add("reportdate", x);

                // report number
                x = tb.Cell((int)MetadataTableRows.finalreviewer + 2, (int)MetadataTableCols.reportdate).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                reportdata.Add("reportnumber", x);

                retval = true;
            }
            else if (tb.Rows.Count == 11 && tb.Columns.Count == 4)
            {
                // document type
                //string x = tb.Rows[(int)MetadataTableRows.doctype].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                //if (!x.ToLower().Contains("central report index (cri) report")) return (retval);
                //reportdata.Add("doctype", "CRI");

                // title
                x = tb.Rows[(int)MetadataTableRows.htitle].Range.Text;
                if (!x.Contains("Title")) return (retval);
                reportdata.Add("title", tb.Rows[(int)MetadataTableRows.title].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty));

                // author
                x = tb.Rows[(int)MetadataTableRows.hauthor].Range.Text;
                if (!x.Contains("Author")) return (retval);
                reportdata.Add("authors", tb.Rows[(int)MetadataTableRows.author].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty));

                // final reviewer
                x = tb.Rows[(int)MetadataTableRows.hfinalreviewer].Range.Text;
                if (!x.Contains("Reviewer")) return (retval);
                reportdata.Add("finalreviewer", tb.Rows[(int)MetadataTableRows.finalreviewer].Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty));

                // report date
                x = tb.Rows[(int)MetadataTableRows.finalreviewer + 1].Range.Text;
                if (!x.Contains("Date")) return (retval);
                x = tb.Cell((int)MetadataTableRows.finalreviewer + 2, (int)MetadataTableCols.elnprj).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                reportdata.Add("reportdate", x);

                // report number
                x = tb.Cell((int)MetadataTableRows.finalreviewer + 2, (int)MetadataTableCols.reportnumber).Range.Text.Trim().Replace("\r", string.Empty).Replace("\a", string.Empty);
                reportdata.Add("reportnumber", x);

                retval = true;
            }
            return (retval);
        }

        internal string PrintMetadata(ICollection<KeyValuePair<string, string>> dt)
        {
            string stsmsg = "";
            // Loop through all the rows in the DataTableReader
            foreach (KeyValuePair<string, string> element in dt)
            {
                stsmsg = string.Concat(stsmsg, element.Key, ": ", element.Value, System.Environment.NewLine);
            }
            return stsmsg;
        }

        enum MetadataTableRows { doctype = 1, htitle, title, hauthor, author, hfinalreviewer, finalreviewer };
        enum MetadataTableCols { finalreviewer = 1, elnprj, reportdate, reportnumber };
    }
}
