﻿using GLAAS.PdfPoc.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Collections.Specialized;

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

                model = GenerateDataMappingOxml(Server.MapPath(string.Format(@"{0}\{1}", _uploadPath, file.FileName)), model);

                //wordDoc(Server.MapPath(string.Format(@"{0}\{1}", _uploadPath, file.FileName)), Server.MapPath(string.Format(@"{0}\{1}", _uploadPath, _templateToDoc)));
                model.FileName = file.FileName;

            }
            catch (Exception ex)
            {
                model.Error = ex.ToString();
            }
            return View(model);
        }

        private WordTemplateModel GenerateDataMappingOxml(string TemplateFileLocation, WordTemplateModel model)
        {
            try
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(TemplateFileLocation, true))
                {

                    // Change the document type to Document
                    document.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

                    // Get the MainPart of the document
                    MainDocumentPart mainPart = document.MainDocumentPart;

                    // Get the Document Settings Part
                    DocumentSettingsPart documentSettingPart1 = mainPart.DocumentSettingsPart;
                    OpenXmlElement[] Enumerate = mainPart.ContentControls().ToArray();
                    List<ModelField> dataMapping = new List<ModelField>();
                    for (int i = 0; i < Enumerate.Count(); i++)
                    {
                        OpenXmlElement cc = Enumerate[i];
                        SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();
                        Tag tag = props.Elements<Tag>().FirstOrDefault();
                        SdtAlias alias = props.Elements<SdtAlias>().FirstOrDefault();
                        string title = ((DocumentFormat.OpenXml.Wordprocessing.StringType)(alias)).Val;
                        string tagName = tag.Val;

                        dataMapping.Add(new ModelField() { Key = title });
                    }

                    model.DataMapping = dataMapping;
                }
            }
            catch (Exception ex)
            {
                throw;
            }

            return model;
        }

        private WordTemplateModel GenerateDataMapping(string TemplateFileLocation, WordTemplateModel model)
        {
            try
            {
                //OBJECT OF MISSING "NULL VALUE"
                Object oMissing = System.Reflection.Missing.Value;
                Object oTemplatePath = TemplateFileLocation;

                Application wordApp = new Application();
                Microsoft.Office.Interop.Word.Document wordDoc = new Microsoft.Office.Interop.Word.Document();

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
                string fileIN = Server.MapPath(string.Format(@"{0}\{1}", _uploadPath, "test123.docx"));
                string fileOUT = fileIN.Replace(".docx", ".pdf");
                var path = model.DocumentType == 0 ? destfilelocation.Replace(".docx", ".pdf") : destfilelocation;


                wordDocOpenXml(tempfilelocation, destfilelocation, model.DataMapping);
                
                if (model.DocumentType == 0)
                {
                    AsposeToPdf(destfilelocation, path);
                }

                //ConvertToPdf(fileIN, fileOUT);
                //wordDoc(tempfilelocation, destfilelocation, model.DataMapping);
                //convertToPdf(destfilelocation);

                var output = new MemoryStream();
                

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
                return View("Index", model);
            }
            catch (Exception ex)
            {
                // return new ContentResult() { Content = ex.Message };
                model.Error = ex.ToString();
                return View("Index", model);
            }

            //return RedirectToAction("Index");
        }

        private void AsposeToPdf(string docXlocation, string pdfLocation)
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(docXlocation);
            doc.Save(pdfLocation, Aspose.Words.SaveFormat.Pdf);
        }

        private void DeleteSdtBlockAndKeepContent(MainDocumentPart mainDocumentPart, string sdtBlockTag)
        {
            List<SdtBlock> sdtList = mainDocumentPart.Document.Descendants<SdtBlock>().ToList();
            SdtBlock sdtA = null;

            foreach (SdtBlock sdt in sdtList)
            {
                if (sdt.SdtProperties.GetFirstChild<Tag>().Val.Value == sdtBlockTag)
                {
                    sdtA = sdt;
                    break;
                }
            }


            OpenXmlElement sdtc = sdtA.GetFirstChild<SdtContentBlock>();
            OpenXmlElement parent = sdtA.Parent;

            OpenXmlElementList elements = sdtc.ChildElements;

            var mySdtc = new SdtContentBlock(sdtc.OuterXml);

            foreach (OpenXmlElement elem in elements)
            {

                string text = parent.FirstChild.InnerText;
                parent.Append((OpenXmlElement)elem.Clone());
            }

            sdtA.Remove();
        }

        public void wordDocOpenXml(string TemplateFileLocation, string GeneratedFileNameLocation, List<ModelField> dataMap)
        {
            if (System.IO.File.Exists(GeneratedFileNameLocation))
            {
                System.IO.File.Delete(GeneratedFileNameLocation);
            }

            System.IO.File.Copy(TemplateFileLocation, GeneratedFileNameLocation);

            using (WordprocessingDocument document = WordprocessingDocument.Open(GeneratedFileNameLocation, true))
            {
                
                // Change the document type to Document
                document.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

                // Get the MainPart of the document
                MainDocumentPart mainPart = document.MainDocumentPart;

                // Get the Document Settings Part
                DocumentSettingsPart documentSettingPart1 = mainPart.DocumentSettingsPart;

                OpenXmlElement[] Enumerate = mainPart.ContentControls().ToArray();

                for (int i = 0; i < Enumerate.Count(); i++)
                {
                    OpenXmlElement cc = Enumerate[i];
                    SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();
                    Tag tag = props.Elements<Tag>().FirstOrDefault();
                    //Console.WriteLine(tag.Val);
                    SdtAlias alias = props.Elements<SdtAlias>().FirstOrDefault();
                    string title = ((DocumentFormat.OpenXml.Wordprocessing.StringType)(alias)).Val;
                    string tagName = tag.Val;

                    //if (dataMap.Any(f => string.Format("{0}Tag", f.Key) == tagName))
                    if (dataMap.Any(f => f.Key == title))
                    {
                        var valkey = dataMap.FirstOrDefault(f => f.Key == title).Value;
                        object val = POCUtil.DictionaryMappedDocPOC[valkey];

                        OpenXmlElement parentElement = cc.Parent;


                        DocumentFormat.OpenXml.Wordprocessing.Paragraph pg = cc.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
                        if (pg != null || true)
                        {
                            //ParagraphProperties paragraphProperties = (ParagraphProperties)cc.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault().ParagraphProperties.Clone();

                            Run r1 = null;
                            DocumentFormat.OpenXml.Wordprocessing.Paragraph p1 = null;

                            if (cc.Parent.GetType() != typeof(DocumentFormat.OpenXml.Wordprocessing.Paragraph))
                            {
                                p1 = parentElement.InsertAfter(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(), cc);
                                r1 = p1.AppendChild(new Run());
                            }
                            else
                            {
                                r1 = parentElement.InsertAfter(new Run(), cc);
                            }

                            r1.RunProperties = new RunProperties();

                            cc.Descendants<DocumentFormat.OpenXml.Wordprocessing.RunProperties>().ToList().ForEach(
                                runProperty => runProperty.ToList().ForEach(
                                    property =>
                                    {
                                        if (!r1.RunProperties.ChildElements.ToList().Exists(propertyToAdd => propertyToAdd.GetType() == property.GetType()))
                                        {
                                            r1.RunProperties.AppendChild((OpenXmlElement)property.CloneNode(true));
                                        }
                                    }
                                     )
                                );

                            if (cc.Descendants<DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox>().Count() > 0)
                            {
                                //☒
                                //☐
                                //☐
                                var t2 = r1.AppendChild(new Text("☒"));
                                //DocumentFormat.OpenXml.Wordprocessing.CheckBox c1 = r1.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.CheckBox());
                                //c1.AppendChild(new Checked());
                            }
                            else
                            {
                                var t1 = r1.AppendChild(new Text(val.ToString()));
                            }
                        }
                    }
                }


                while (mainPart.ContentControls().Count() > 0)
                {
                    mainPart.ContentControls().FirstOrDefault().Remove();
                }
                
                // Save the document
                mainPart.Document.Save();
                document.Close();
            }
        }

        public void wordDoc(string TemplateFileLocation, string GeneratedFileNameLocation, List<ModelField> dataMap)
        {
            Application wordApp = new Application();
            Microsoft.Office.Interop.Word.Document wordDoc = new Microsoft.Office.Interop.Word.Document();
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
                            wordApp.Selection.TypeText(POCUtil.GenerateLoremIpsum());
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

        private void convertToPdf(string docLocation)
        {
            //OBJECT OF MISSING "NULL VALUE"
            Object oMissing = System.Reflection.Missing.Value;
            Application wordApp = new Application();
            Microsoft.Office.Interop.Word.Document wordDoc = null;
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

        private void ConvertToPdf(string fileIN, string fileOUT)
        {
            //string projectDir = Server.MapPath("~/");

            NameValueCollection commonLoggingproperties = new NameValueCollection();
            commonLoggingproperties["showDateTime"] = "false";
            commonLoggingproperties["level"] = "INFO";
            Common.Logging.LogManager.Adapter = new Common.Logging.Simple.ConsoleOutLoggerFactoryAdapter(commonLoggingproperties);


            Common.Logging.ILog log = Common.Logging.LogManager.GetCurrentClassLogger();
            log.Info("Hello from Common Logging");

            // Necessary, if slf4j-api and slf4j-NetCommonLogging are separate DLLs
            ikvm.runtime.Startup.addBootClassPathAssembly(
                System.Reflection.Assembly.GetAssembly(
                    typeof(org.slf4j.impl.StaticLoggerBinder)));

            // Configure to find docx4j.properties
            // .. add as URL the dir containing docx4j.properties (not the file itself!)
            Plutext.PropertiesConfigurator.setDocx4jPropertiesDir(Server.MapPath("~/src/samples/resources/"));

            //org.docx4j.openpackaging.parts.WordprocessingML.ObfuscatedFontPart.getTemporaryEmbeddedFontsDir()

            // OK, do it..
            org.docx4j.openpackaging.packages.WordprocessingMLPackage wordMLPackage = org.docx4j.openpackaging.packages.WordprocessingMLPackage.load(new java.io.File(fileIN));

            java.io.FileOutputStream fos = new java.io.FileOutputStream(new java.io.File(fileOUT));

            org.docx4j.Docx4J.toPDF(wordMLPackage, fos);

            fos.close();
        }
    }
}