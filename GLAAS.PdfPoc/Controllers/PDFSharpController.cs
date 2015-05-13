using GLAAS.PdfPoc.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GLAAS.PdfPoc.Controllers
{
    public class PDFSharpController : Controller
    {
        private string _uploadPath = "~/PDFForms/";

        // GET: Home
        public ActionResult Index()
        {
            return View(new UploadModel());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Index(UploadModel model)
        {
            try
            {
                HttpPostedFileBase file = model.File;
                
                var pdfDocument = PdfReader.Open(file.InputStream);
                
                file.SaveAs(Server.MapPath(_uploadPath + file.FileName));
                model.FileName = file.FileName;

                foreach(PdfTextField field in pdfDocument.AcroForm.Fields)
                {
                    model.Fields.Add(new Field { Key = field.Name, Value = string.Empty });
                }
            }
            catch(Exception ex)
            {
                
            }
            return View(model);
        }

        [HttpPost]
        public ActionResult Generate(UploadModel model)
        {
            try
            {
                var pdfDocument = PdfReader.Open(Server.MapPath(_uploadPath + model.FileName));
                var output = new MemoryStream();

                foreach(var field in model.Fields)
                {
                    (pdfDocument.AcroForm.Fields[field.Key] as PdfTextField).Text = field.Value;
                }
                
                pdfDocument.Save(output);
                pdfDocument.Dispose();
                
                Response.AddHeader("Content-Disposition", "attachment; filename=GeneratedFrom-" + (model.FileName.ToLower().EndsWith(".pdf") ? model.FileName : model.FileName + ".pdf"));
                Response.ContentType = "application/pdf";
                Response.BinaryWrite(output.ToArray());
                Response.End();
            }
            catch (Exception ex)
            {

            }

            return RedirectToAction("Index");
        }
    }
}