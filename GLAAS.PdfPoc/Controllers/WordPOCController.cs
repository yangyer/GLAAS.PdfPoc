using GLAAS.PdfPoc.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GLAAS.PdfPoc.Controllers
{
    public class WordPOCController : Controller
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
                //model.FileName = file.FileName;

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

                //Dictionary<int, string> DataDictionaryReplacementValues = new Dictionary<int, string>();
                ////YourName
                //DataDictionaryReplacementValues[0] = "Mohammad Murtaza Zaidi";
                //// bool Single
                //DataDictionaryReplacementValues[0] = "true";


                //var pdfReader = new PdfReader(Server.MapPath(_uploadPath + model.FileName));
                //var output = new MemoryStream();
                //var stamper = new PdfStamper(pdfReader, output);

                //foreach (var field in model.Fields)
                //{
                //    stamper.AcroFields.SetField(field.Key, field.Value);
                //}

                //stamper.FormFlattening = true;
                //stamper.Close();
                //pdfReader.Close();



                //Response.AddHeader("Content-Disposition", "attachment; filename=GeneratedFrom-" + (model.FileName.ToLower().EndsWith(".pdf") ? model.FileName : model.FileName + ".pdf"));
                //Response.ContentType = "application/pdf";
                //Response.BinaryWrite(output.ToArray());
                //Response.End();
            }
            catch (Exception ex)
            {

            }

            return RedirectToAction("Index");
        }

        public void wordDoc(string TemplateFileLocation, string GeneratedFileNameLocation)
        {
            try
            {
                //OBJECT OF MISSING "NULL VALUE"
                Object oMissing = System.Reflection.Missing.Value;
                Object oTemplatePath = TemplateFileLocation;

                Application wordApp = new Application();
                Document wordDoc = new Document();

                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

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

    }
}