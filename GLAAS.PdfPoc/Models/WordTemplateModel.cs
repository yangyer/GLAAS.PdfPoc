using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


namespace GLAAS.PdfPoc.Models
{
    public class WordTemplateModel
    {
        public HttpPostedFileBase File { get; set; }
        public string FileName { get; set; }

        public Dictionary<string, string> DataDictionary
        {
            get
            {
                return GLAAS.PdfPoc.POCUtil.DataDictionary;
            }
        }

        //public Dictionary<string, string> DataMapping { get; set; }

        public List<ModelField> DataMapping { get; set; }

    }

    public class ModelField
    {
        public string Key { get; set; }
        public string Value { get; set; }

        public ModelField()
        {

        }
    }
}