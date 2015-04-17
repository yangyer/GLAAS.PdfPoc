using GLAAS.PdfPoc.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GLAAS.PdfPoc.Controllers
{
    public class DocPOCController : Controller
    {

        private string _uploadPath = "~/PDFForms/";
        private string _templateDoc = "templateDoc.dotx";
        private string _templateToDoc = "TemplateToDoc.docx";
        // GET: Home
        public ActionResult Index()
        {
            return View(new WordTemplateModel());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Index(WordTemplateModel model)
        {
            //GLAAS.PdfPoc.POCUtil
            try
            {
                HttpPostedFileBase file = model.File;

                // Save File to Server.
                file.SaveAs(Server.MapPath(_uploadPath + file.FileName));
                // Give user the mapping

                model = GenerateDataMapping(Server.MapPath(string.Format(@"{0}\{1}", _uploadPath, file.FileName)), model);

                //wordDoc(Server.MapPath(string.Format(@"{0}\{1}", _uploadPath, file.FileName)), Server.MapPath(string.Format(@"{0}\{1}", _uploadPath, _templateToDoc)));
                model.FileName = file.FileName;

            }
            catch (Exception ex)
            {

            }
            return View(model);
        }

        private WordTemplateModel GenerateDataMapping(string TemplateFileLocation, WordTemplateModel model)
        {
            try
            {
                //OBJECT OF MISSING "NULL VALUE"
                Object oMissing = System.Reflection.Missing.Value;
                Object oTemplatePath = TemplateFileLocation;

                Application wordApp = new Application();
                Document wordDoc = new Document();

                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                //Dictionary<string, string> DataMapping = new Dictionary<string, string>();
                List<ModelField> dataMapping = new List<ModelField>();

                foreach (Microsoft.Office.Interop.Word.ContentControl cc in wordDoc.ContentControls)
                {
                    //DataMapping[cc.Title] = "";
                    dataMapping.Add(new ModelField() { Key = cc.Title });
                }
                model.DataMapping = dataMapping;

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

            return model;
        }

        [HttpPost]
        public ActionResult Generate(WordTemplateModel model)
        {
            try
            {
                string tempfilelocation = Server.MapPath(string.Format(@"{0}\{1}", _uploadPath, model.FileName));
                string destfilelocation = Server.MapPath(string.Format(@"{0}\{1}", _uploadPath, _templateToDoc));
                wordDoc(tempfilelocation, destfilelocation, model.DataMapping);
                convertToPdf(destfilelocation);
                
                var output = new MemoryStream();
                var path = model.DocumentType == 0 ? destfilelocation.Replace(".docx", ".pdf") : destfilelocation;

                using (var fsSource = new System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    byte[] bytes = new byte[fsSource.Length];
                    int numBytesToRead = (int)fsSource.Length;
                    int numBytesRead = 0;
                    while (numBytesToRead > 0)
                    {
                        // Read may return anything from 0 to numBytesToRead. 
                        int n = fsSource.Read(bytes, numBytesRead, numBytesToRead);

                        // Break when the end of the file is reached. 
                        if (n == 0)
                            break;

                        numBytesRead += n;
                        numBytesToRead -= n;
                    }
                    numBytesToRead = bytes.Length;

                    output.Write(bytes, 0, numBytesToRead);
                }


                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=GeneratedFrom-{0}", model.DocumentType == 0 ? "GeneratedDocument.pdf" : "GeneratedDocument.docx"));
                Response.ContentType = "application/pdf";
                Response.BinaryWrite(output.ToArray());
                Response.End();
            }
            catch (Exception ex)
            {

            }

            return RedirectToAction("Index");
        }

        public void wordDoc(string TemplateFileLocation, string GeneratedFileNameLocation, List<ModelField> dataMap)
        {
            Application wordApp = new Application();
            Document wordDoc = new Document();
            //OBJECT OF MISSING "NULL VALUE"
            Object oMissing = System.Reflection.Missing.Value;
            Object oTemplatePath = TemplateFileLocation;

            try
            {


                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                foreach (Microsoft.Office.Interop.Word.ContentControl cc in wordDoc.ContentControls)
                {
                    string fieldName = cc.Title;

                    if (dataMap.Any(f => f.Key == fieldName))
                    {
                        var valkey = dataMap.FirstOrDefault(f => f.Key == fieldName).Value;
                        object val = POCUtil.DictionaryMappedDocPOC[valkey];
                        if (cc.Type == WdContentControlType.wdContentControlCheckBox)
                        {
                            cc.Checked = (bool)val;
                        }
                        else if (cc.Type == WdContentControlType.wdContentControlText)
                        {
                            cc.Range.Text = val.ToString();
                        }
                    }
                }

                foreach (Microsoft.Office.Interop.Word.Field myMergeField in wordDoc.Fields)
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
                            wordApp.Selection.TypeText(GenerateLoremIpsum());
                        }
                    }
                }

                wordDoc.SaveAs(GeneratedFileNameLocation);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                object doNotSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)wordDoc).Close(ref doNotSaveChanges, ref oMissing, ref oMissing);
                wordDoc = null;
                ((_Application)wordApp).Quit(ref doNotSaveChanges, ref oMissing, ref oMissing);
                wordApp = null;
            }

        }

        private string GenerateLoremIpsum()
        {
            return @"Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
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
                    desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.";
        }

        private void convertToPdf(string docLocation)
        {
            //OBJECT OF MISSING "NULL VALUE"
            Object oMissing = System.Reflection.Missing.Value;
            Application wordApp = new Application();
            Document wordDoc = null;
            try
            {
                wordDoc = wordApp.Documents.Open(docLocation);


                //wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                object outputFileName = wordDoc.FullName.Replace(".docx", ".pdf");
                object fileFormat = WdSaveFormat.wdFormatPDF;
                wordDoc.SaveAs(ref outputFileName,
                                ref fileFormat, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                //wordApp.Documents.Open("myFile.doc");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                object doNotSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                if (wordDoc != null)
                {
                    ((_Document)wordDoc).Close(ref doNotSaveChanges, ref oMissing, ref oMissing);
                }
                wordDoc = null;
                ((_Application)wordApp).Quit(ref doNotSaveChanges, ref oMissing, ref oMissing);
                wordApp = null;
            }

            //OpenOffice o = new OpenOffice();
            //Console.WriteLine(o.ExportToPdf("C:\\MyProjects\\myfile.docx").ToString());
        }
    }
}