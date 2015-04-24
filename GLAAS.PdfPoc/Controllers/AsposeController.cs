using Aspose.Words;
using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GLAAS.PdfPoc.Controllers
{
    public class AsposeController : Controller
    {
        private string _uploadPath = "~/PDFForms/";
        private string _myfile = "myfile.docx";
        private string _templateToDoc = "TemplateToDoc.docx";

        // GET: Aspose
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ConvertToPdf()
        {
            string docxFilePath = Server.MapPath(string.Format("{0}/{1}", _uploadPath, _myfile));
            string pdfFilePath = docxFilePath.Replace("docx", "pdf");
            string docxFilePath2 = Server.MapPath(string.Format("{0}/{1}", _uploadPath, _templateToDoc));

            if (System.IO.File.Exists(pdfFilePath))
            {
                System.IO.File.Delete(pdfFilePath);
            }

            MemoryStream output = new MemoryStream();
            MemoryStream output2 = new MemoryStream();
            MemoryStream output3 = new MemoryStream();

            Document doc = new Document(docxFilePath2);
            Document doc2 = new Document(docxFilePath2);


            doc.Save(output, SaveFormat.Pdf);
            doc2.Save(output2, SaveFormat.Pdf);

            Aspose.Pdf.Document pdf1 = new Aspose.Pdf.Document(output);
            Aspose.Pdf.Document pdf2 = new Aspose.Pdf.Document(output2);

            pdf1.Pages.Add(pdf2.Pages);

            pdf1.Save(output3);

            Response.AddHeader("Content-Disposition", string.Format("attachment; filename=output.pdf"));
            Response.ContentType = "application/pdf";
            Response.BinaryWrite(output3.ToArray());
            Response.End();

            return View("Index");
        }
    }
}