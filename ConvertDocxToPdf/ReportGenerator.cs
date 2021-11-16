using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using ConvertDocxToPdf.DocXToPdfHandlers;

namespace ConvertDocxToPdf
{
    public class ReportGenerator
    {
        private readonly string _locationOfLibreOfficeSoffice;

        public ReportGenerator(string locationOfLibreOfficeSoffice = "")
        {
            _locationOfLibreOfficeSoffice = locationOfLibreOfficeSoffice;
        }
        
        public void Convert(string inputFile, string outputFile, Placeholders rep= null)
        {
                if (outputFile.EndsWith(".docx"))
                {
                    GenerateReportFromDocxToDocX(inputFile, outputFile, rep);
                }
                else if(outputFile.EndsWith(".pdf"))
                {
                    GenerateReportFromDocxToPdf(inputFile, outputFile, rep);
                }
                else if (outputFile.EndsWith(".html") || outputFile.EndsWith(".htm"))
                {
                    GenerateReportFromDocxToHtml(inputFile, outputFile, rep);
                }
        }

        public void Print(string templateFile, string printerName = null, Placeholders rep=null)
        {
            if (rep != null)
            {
                if (templateFile.EndsWith(".docx"))
                {
                    PrintDocx(templateFile, printerName, rep);
                }
                else if (templateFile.EndsWith(".html") || templateFile.EndsWith(".htm"))
                {
                    PrintHtml(templateFile, printerName, rep);
                }
                else
                {
                    LibreOfficeWrapper.Print(templateFile, printerName, _locationOfLibreOfficeSoffice);
                }
            }
            else
            {
                LibreOfficeWrapper.Print(templateFile, printerName, _locationOfLibreOfficeSoffice);
            }
        }

        private void PrintDocx(string templateFile, string printername, Placeholders rep)
        {
            var docx = new DocXHandler(templateFile, rep);
            var ms = docx.ReplaceAll();
            var tempFileToPrint = Path.ChangeExtension(Path.GetTempFileName(), ".docx");
            StreamHandler.WriteMemoryStreamToDisk(ms, tempFileToPrint);
            LibraOfficeWrapper.Print(tempFileToPrint, printername, _locationOfLibreOfficeSoffice);
            File.Delete(tempFileToPrint);
        }

        private void PrintDocx(string templateFile, string printername, Placeholders rep)
        {
            var htmlContent = File.ReadAllText(templateFile);
            htmlContent = HtmlHandler.ReplaceAll(htmlContent, rep);
            var tempFileToPrint = Path.ChangeExtension(Path.GetTempFileName(), ".html");
            File.WriteAllText(tempFileToPrint, htmlContent);
            LibraOfficeWrapper.Print(tempFileToPrint, printername, _locationOfLibreOfficeSoffice);
            File.Delete(tempFileToPrint);
        }

        private void GenerateReportFromDocxToDocX(string docxSource, string docxTarget, Placeholders rep)
        {
            var docx = new DocXHandler(docxSource, rep);
            var ms = docx.ReplaceAll();
            StreamHandler.WriteMemoryStreamToDisk(ms, docxTarget);
        }

        private void GenerateReportFromDocxToPdf(string docxSource, string pdfTarget, Placeholders rep)
        {
            var docx = new DocXHandler(docxSource, rep);
            var ms = docx.ReplaceAll();
            var tmpFile = Path.Combine(Path.GetDirectoryName(pdfTarget), Path.GetFileNameWithoutExtension(pdfTarget) + Guid.NewGuid().ToString().Substring(0, 10) + ".docx");
            StreamHandler.WriteMemoryStreamToDisk(ms, tmpFile);
            LibraOfficeWrapper.Convert(tmpFile, pdfTarget, _locationOfLibreOfficeSoffice);
            File.Delete(tmpFile);
        }

    }
}
