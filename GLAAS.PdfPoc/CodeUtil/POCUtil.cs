using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
                d["2"] = "RequestAppointMentTypeNew";
                d["3"] = "EmployeeFullName";
                d["4"] = "EmployeePrimaryEmailAddress";
                d["5"] = "SchoolEmailAddress";
                d["6"] = "AppointmentType";
                d["7"] = "Semester";
                d["8"] = "PayStartDate";
                d["9"] = "PayEndDate";
                d["10"] = "Percentage";
                d["11"] = "MaxHours";
                d["12"] = "TotalSalary";
                d["13"] = "AnnualBaseSalary";
                d["14"] = "EmployeeFirstName";
                d["15"] = "EmployeeLastName";
                d["16"] = "RequestAppointMentTypeExtension";
                d["17"] = "MultiLineContent";
                return d;
            }
        }

        public static Dictionary<string, object> DictionaryMapped
        {
            get
            {
                Dictionary<string, object> d = new Dictionary<string, object>();
                d["0"] = "Mohammad Zaidi";
                d["1"] = true;
                return d;
            }
        }


        public static Dictionary<string, object> DictionaryMappedDocPOC
        {
            get
            {
                Dictionary<string, object> d = new Dictionary<string, object>();
                d["0"] = null;
                d["1"] = null;
                d["2"] = "New Appointment";
                d["3"] = "John Doe";
                d["4"] = "Joe@ucmerced.edu";
                d["5"] = "eng@ucmerced.edu";
                d["6"] = "Teaching Assistant (TA)";
                d["7"] = "Fall 2015";
                d["8"] = "7/1/2015";
                d["9"] = "12/31/2015";
                d["10"] = "25";
                d["11"] = "40";
                d["12"] = "$70,000";
                d["13"] = "$85,000";
                d["14"] = "John";
                d["15"] = "Doe";
                d["16"] = null;
                d["17"] = GenerateLoremIpsum();
                return d;
            }
        }

        public static string GenerateLoremIpsum()
        {
            return @"Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, 
                    when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with 
                    desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.";
        }
    }

    public static class ContentControlExtensions
    {
        public static IEnumerable<OpenXmlElement> ContentControls(this OpenXmlPart part)
        {
            return part.RootElement.Descendants().Where(e => e is SdtBlock || e is SdtRun || e is SdtCell);
        }

        public static IEnumerable<OpenXmlElement> ContentControls(this WordprocessingDocument doc)
        {
            foreach (var cc in doc.MainDocumentPart.ContentControls())
                yield return cc;
            foreach (var header in doc.MainDocumentPart.HeaderParts)
                foreach (var cc in header.ContentControls())
                    yield return cc;
            foreach (var footer in doc.MainDocumentPart.FooterParts)
                foreach (var cc in footer.ContentControls())
                    yield return cc;
            if (doc.MainDocumentPart.FootnotesPart != null)
                foreach (var cc in doc.MainDocumentPart.FootnotesPart.ContentControls())
                    yield return cc;
            if (doc.MainDocumentPart.EndnotesPart != null)
                foreach (var cc in doc.MainDocumentPart.EndnotesPart.ContentControls())
                    yield return cc;
        }

       
    }
}