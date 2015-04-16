using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GLAAS.PdfPoc
{
    public static class POCUtil
    {
        public static Dictionary<string, string> DataDictionary
        {
            get
            {
                Dictionary<string, string> d = new Dictionary<string, string>();
                d["0"] = "StudentName";
                d["1"] = "SingleOrMarried";
                return d;
            }
        }

        public static Dictionary<string, object> DictionaryMapped
        {
            get
            {
                Dictionary<string, object> d = new Dictionary<string, object>();
                d["0"] = "Zaidi";
                d["1"] = true;
                return d;
            }
        }


    }
}