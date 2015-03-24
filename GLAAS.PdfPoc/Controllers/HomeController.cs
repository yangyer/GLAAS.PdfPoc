using GLAAS.PdfPoc.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GLAAS.PdfPoc.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Index(UploadModel model)
        {
            try
            {
                HttpPostedFileBase file = model.File;


            }
            catch(Exception ex)
            {
                
            }
            return View();
        }
    }
}