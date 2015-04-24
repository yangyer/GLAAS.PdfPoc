using PDFTech;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GLAAS.PdfPoc.Controllers
{
    public class AcroController : Controller
    {
        private string _uploadPath = "~/PDFForms/";
        private string _myfile = "PDFTemplate.pdf";
        private string _templateToDoc = "Generated.pdf";

        // GET: Aspose
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ConvertToPdf()
        {
            string templateFilePath = Server.MapPath(string.Format("{0}/{1}", _uploadPath, _myfile));
            string documentFilePath = Server.MapPath(string.Format("{0}/{1}", _uploadPath, _templateToDoc));

            if (System.IO.File.Exists(documentFilePath))
            {
                System.IO.File.Delete(documentFilePath);
            }

            MemoryStream output = new MemoryStream();

            PDFDocument doc = new PDFDocument(documentFilePath);

            doc.LoadPdf(templateFilePath, "");

            for (int i = 0; i < doc.AcroForm.Fields.Count; i++)
            {
                var c = doc.AcroForm.Fields[i];
                if(POCUtil.DataDictionary.Values.Contains(c.Name))
                {
                    //get the key
                    var valkey = POCUtil.DataDictionary.FirstOrDefault(f => f.Value == c.Name).Key;
                    // get the mapped value
                    object val = POCUtil.DictionaryMappedDocPOC[valkey];

                    if(c.Name == "SingleOrMarried")
                    {
                        (c as PDFCheckBox).Checked = true;
                    }
                    else
                    {
                        (c as PDFEdit).Text = val.ToString();
                    }
                }
                else
                {
                    switch (c.Name)
                    {
                        case "RequestAppointMentType":
                            (c as PDFEdit).Text = POCUtil.DictionaryMappedDocPOC["2"].ToString();
                            break;
                        case "EmailedTo":
                            (c as PDFEdit).Text = POCUtil.DictionaryMappedDocPOC["4"].ToString();
                            break;
                    }
                }
            }

            doc.Save();

            using (var fsSource = new System.IO.FileStream(documentFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
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

            Response.AddHeader("Content-Disposition", string.Format("attachment; filename=output.pdf"));
            Response.ContentType = "application/pdf";
            Response.BinaryWrite(output.ToArray());
            Response.End();

            return View("Index");
        }
    }
}