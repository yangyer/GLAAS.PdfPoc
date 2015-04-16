using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Novacode;
using System.Diagnostics;
using QOpenOffice;
using System.IO;
namespace TestConsole1
{
    class Program
    {
        public static string FileSaveAs
        {
            get { return @"C:\MyProjects\myfile.docx"; }
        }


        static void Main(string[] args)
        {
            try
            {
                //testWord();
                wordDoc();
                //convertToPdf();
                //fillPdf();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static void convertToPdf()
        {
            try
            {
                //OBJECT OF MISSING "NULL VALUE"

                Object oMissing = System.Reflection.Missing.Value;

                Application wordApp = new Application();
                Document wordDoc = wordApp.Documents.Open(FileSaveAs);

                //wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                object outputFileName = wordDoc.FullName.Replace(".docx", ".pdf");
                object fileFormat = WdSaveFormat.wdFormatPDF;
                wordDoc.SaveAs(ref outputFileName,
                                ref fileFormat, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                //wordApp.Documents.Open("myFile.doc");
                object doNotSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)wordDoc).Close(ref doNotSaveChanges, ref oMissing, ref oMissing);
                wordDoc = null;
                ((_Application)wordApp).Quit(ref doNotSaveChanges, ref oMissing, ref oMissing);
                wordApp = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            //OpenOffice o = new OpenOffice();
            //Console.WriteLine(o.ExportToPdf("C:\\MyProjects\\myfile.docx").ToString());
        }

        private static void testWord()
        {
            try
            {
                //OBJECT OF MISSING "NULL VALUE"

                Object oMissing = System.Reflection.Missing.Value;

                Application wordApp = new Application();
                Document wordDoc = wordApp.Documents.Open(FileSaveAs);

                //foreach (Microsoft.Office.Interop.Word.FormField field in wordDoc.FormFields)
                //{
                //    switch (field.Name)
                //    {
                //        case "MyName":
                //            field.Range.Text = "Mohammad Murtaza Zaidi";
                //            break;

                //        default:
                //            break;
                //    }
                //}

                foreach (Microsoft.Office.Interop.Word.ContentControl cc in wordDoc.ContentControls)
                {

                    switch (cc.Title)
                    {
                        case "MyName":
                            cc.Range.Text = "Mohammad Murtaza Zaidi";
                            break;

                        default:
                            break;
                    }
                   
                }
                wordDoc.Save();

                object doNotSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)wordDoc).Close(ref doNotSaveChanges, ref oMissing, ref oMissing);
                wordDoc = null;
                ((_Application)wordDoc).Quit(ref doNotSaveChanges, ref oMissing, ref oMissing);
                wordApp = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            //OpenOffice o = new OpenOffice();
            //Console.WriteLine(o.ExportToPdf("C:\\MyProjects\\myfile.docx").ToString());
        }


        private static void fillPdf()
        {
            try
            {
                //OBJECT OF MISSING "NULL VALUE"

                Object oMissing = System.Reflection.Missing.Value;

                Application wordApp = new Application();
                Document pdfDoc = wordApp.Documents.Open(@"E:\Git\Repos\GLAAS.PdfPoc\Misc\test1.pdf");

                //wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                //object outputFileName = pdfDoc.FullName.Replace(".docx", ".pdf");
                object fileFormat = WdSaveFormat.wdFormatPDF;
                //wordDoc.SaveAs(ref outputFileName,
                //                ref fileFormat, ref oMissing, ref oMissing,
                //                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                //                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                //                ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                //wordApp.Documents.Open("myFile.doc");
                object doNotSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)pdfDoc).Close(ref doNotSaveChanges, ref oMissing, ref oMissing);
                pdfDoc = null;
                ((_Application)pdfDoc).Quit(ref doNotSaveChanges, ref oMissing, ref oMissing);
                wordApp = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            //OpenOffice o = new OpenOffice();
            //Console.WriteLine(o.ExportToPdf("C:\\MyProjects\\myfile.docx").ToString());
        }

        public static void docx()
        {

            // Modify to suit your machine:
            string fileName = @"C:\MyProjects\DocXExample.docx";

            //DocX d = new DocX();


            // Create a document in memory:
            var doc = DocX.Create(fileName);

            // Insert a paragrpah:
            doc.InsertParagraph("This is my first paragraph");

            // Save to the output directory:
            doc.Save();

            // Open in Word:
            Process.Start("WINWORD.EXE", fileName);
        }

        public static void wordDoc()
        {
            try
            {
                //OBJECT OF MISSING "NULL VALUE"

                Object oMissing = System.Reflection.Missing.Value;

                Object oTemplatePath = "C:\\MyProjects\\templateDoc.dotx";


                Application wordApp = new Application();
                Document wordDoc = new Document();

                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);


                //foreach (Microsoft.Office.Interop.Word.FormField field in wordDoc.FormFields)
                //{
                //    switch (field.Name)
                //    {
                //        case "MyName":
                //            field.Range.Text = "Mohammad Murtaza Zaidi";
                //            break;

                //        default:
                //            break;
                //    }
                //}

                foreach (Microsoft.Office.Interop.Word.ContentControl cc in wordDoc.ContentControls)
                {

                    switch (cc.Title)
                    {
                        case "MyName":
                            cc.Range.Text = "Mohammad Murtaza Zaidi";
                            break;
                        case "Single":
                            cc.Checked = true;
                            break;
                        default:
                            break;
                    }

                }

                foreach (Field myMergeField in wordDoc.Fields)
                {


                    Range rngFieldCode = myMergeField.Code;

                    String fieldText = rngFieldCode.Text;



                    // ONLY GETTING THE MAILMERGE FIELDS

                    if (fieldText.StartsWith(" MERGEFIELD"))
                    {

                        // THE TEXT COMES IN THE FORMAT OF

                        // MERGEFIELD  MyFieldName  \\* MERGEFORMAT

                        // THIS HAS TO BE EDITED TO GET ONLY THE FIELDNAME "MyFieldName"

                        Int32 endMerge = fieldText.IndexOf("\\");

                        Int32 fieldNameLength = fieldText.Length - endMerge;

                        String fieldName = fieldText.Substring(11, endMerge - 11);

                        // GIVES THE FIELDNAMES AS THE USER HAD ENTERED IN .dot FILE

                        fieldName = fieldName.Trim();

                        // **** FIELD REPLACEMENT IMPLEMENTATION GOES HERE ****//

                        // THE PROGRAMMER CAN HAVE HIS OWN IMPLEMENTATIONS HERE
                        if (fieldName == "StudentName")
                        {
                            myMergeField.Select();

                            wordApp.Selection.TypeText(@"Yer Yang");
                        }
                        if (fieldName == "DocumentBody")
                        {

                            myMergeField.Select();

                            wordApp.Selection.TypeText(@"
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.");

                        }

                    }

                }
                wordDoc.SaveAs("C:\\MyProjects\\myfile.docx");
                //wordApp.Documents.Open("myFile.doc");
                object doNotSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)wordDoc).Close(ref doNotSaveChanges, ref oMissing, ref oMissing);
                wordDoc = null;
                ((_Application)wordApp).Quit(ref doNotSaveChanges, ref oMissing, ref oMissing);
                wordApp = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


    }
}
