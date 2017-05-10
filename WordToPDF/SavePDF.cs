using System;
using System.Diagnostics;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace WordToPDF
{
    public class SavePDF
    {
        public bool Save(Word.Document Doc, string pdffullfilename)
        {
            bool retval = false;
            try
            {
                string docfilename = Doc.Name;
                //string filedir = Doc.Path;

                //string pdffilename = docfilename.Replace(Path.GetExtension(docfilename), ".pdf");
                Word.WdExportFormat exportFormat = Word.WdExportFormat.wdExportFormatPDF;
                Word.WdExportOptimizeFor exportOptimizeFor = Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
                Word.WdExportRange exportRange = Word.WdExportRange.wdExportAllDocument;
                //Word.WdExportItem exportItem = Word.WdExportItem.wdExportDocumentWithMarkup;
                Word.WdExportItem exportItem = Word.WdExportItem.wdExportDocumentContent;
                Word.WdExportCreateBookmarks createBookmarks = Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;

                bool includeDocProps = true;
                bool keepIRM = true;
                bool docStructureTags = true;
                bool bitmapMissingFonts = true;
                bool useISO19005_1 = true;
                bool openAfterExport = false;

                //If Export Range Is Set
                int startPage = 0;
                int endPage = 0;

                object missing = Missing.Value;

                // Export it in the specified format.  
                if (Doc != null)
                    Doc.ExportAsFixedFormat(
                        pdffullfilename /*Path.Combine(filedir, pdffilename)*/,
                        exportFormat,
                        openAfterExport,
                        exportOptimizeFor,
                        exportRange,
                        startPage,
                        endPage,
                        exportItem,
                        includeDocProps,
                        keepIRM,
                        createBookmarks,
                        docStructureTags,
                        bitmapMissingFonts,
                        useISO19005_1,
                        ref missing);
                retval = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("saveaspdf: " + ex.ToString());
                throw (ex);
            }
            return (retval);
        }
    }
}
