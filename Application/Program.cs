using System.Collections.Generic;
using System.IO;
using System.Reflection;
using DocXToPdfConverter;
using DocXToPdfConverter.DocXToPdfHandlers;

namespace ConvertDocxToPdf
{
    public static class Program
    {

        public static void Main(string[] args)
        {
            string locationOfLibreOfficeSoffice =
               @"C:\Program Files\LibreOffice\program\soffice.exe";


            string executableLocation = Path.GetDirectoryName(
                           Assembly.GetExecutingAssembly().Location);

            string docxLocate = Path.Combine(executableLocation, "Hidayu.docx");

            var placeholders = new Placeholders
            {
                NewLineTag = "<br/>",
                TextPlaceholderStartTag = "##",
                TextPlaceholderEndTag = "##",
                TablePlaceholderStartTag = "==",
                TablePlaceholderEndTag = "==",

                TextPlaceholders = new Dictionary<string, string>
            {
                {"Company", "VirtualX"},
                {"Street", " 21-3, JALAN 2/109F, TAMAN DANAU KOTA,"},
                {"City", "FEDERAL TERITORY OF KUALA LUMPUR"},
                {"Website", "www.virtualx.my"},
                {"DateTest", "8 November 2021"}
            },

                HyperlinkPlaceholders = new Dictionary<string, HyperlinkElement>
                {
                    {"Website", new HyperlinkElement{ Link= "http://www.virtualx.my", Text="www.virtualx.my" } }
                },

                TablePlaceholders = new Dictionary<string, string>
                {

                        {"Name", "Hidayu"},
                        {"TrainingMode", "AR Training"},

                }
            };

            var test = new ReportGenerator(locationOfLibreOfficeSoffice);
            
            test.Convert(docxLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Hidayu.pdf"), placeholders);

        }
    }
}
