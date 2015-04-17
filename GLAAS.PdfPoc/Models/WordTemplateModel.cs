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

        public List<ModelField> DocumentTypes
        {
            get
            { 
                List<ModelField> m = new List<ModelField>();
                m.Add(new ModelField() { Key = "0", Value = "PDF Document" });
                m.Add(new ModelField() { Key = "1", Value = "Word Document" });
                return m;
            }
        }
        public int DocumentType { get; set; }
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