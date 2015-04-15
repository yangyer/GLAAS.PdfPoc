using GLAAS.PdfPoc.Models;
using iTextSharp.text.pdf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GLAAS.PdfPoc.Controllers
{
    public class HomeController : Controller
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
                
                var pdfReader = new PdfReader(file.InputStream);
                //var output = new MemoryStream();
                //var stamper = new PdfStamper(pdfReader, output);
                file.SaveAs(Server.MapPath(_uploadPath + file.FileName));
                model.FileName = file.FileName;
                foreach(DictionaryEntry field in pdfReader.AcroFields.Fields)
                {
                    model.Fields.Add(new Field { Key = field.Key.ToString(), Value = string.Empty });
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
                var pdfReader = new PdfReader(Server.MapPath(_uploadPath + model.FileName));
                var output = new MemoryStream();
                var stamper = new PdfStamper(pdfReader, output);

                foreach(var field in model.Fields)
                {
                    stamper.AcroFields.SetField(field.Key, field.Value);
                }

                stamper.FormFlattening = true;
                stamper.Close();
                pdfReader.Close();

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