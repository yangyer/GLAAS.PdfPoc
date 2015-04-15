using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GLAAS.PdfPoc.Models
{
    public class UploadModel
    {
        public HttpPostedFileBase File { get; set; }

        public string FileName { get; set; }

        public List<Field> Fields { get; set; }

        public UploadModel()
        {
            Fields = new List<Field>();
        }
    }

    public class Field
    {
        public string Key { get; set; }
        public string Value { get; set; }

        public Field()
        {

        }
    }
}