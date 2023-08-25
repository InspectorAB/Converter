using System;
using System.IO;
using System.Reflection.Metadata;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using static System.Net.Mime.MediaTypeNames;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace HTMLToWordConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            string htmlFilePath = "D:\\github\\Sorting\\HTMLToWord\\HTMLToWord\\HTMLDocument.html";
            string outputDirectory = "D:\\github\\Sorting\\HTMLToWord\\HTMLToWord"; // Specify the directory where you want to save the Word document

            // Read HTML content
            string htmlContent = File.ReadAllText(htmlFilePath);

            //removes the HTML tag using regex
            string plainTextContent = System.Text.RegularExpressions.Regex.Replace(htmlContent, "<.*?>", "");

            // File generated based on TimeStamp
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string wordFilePath = Path.Combine(outputDirectory, $"output_{timestamp}.docx");

            // Setting content for the word document
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(wordFilePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Convert HTML content to Word elements
                Paragraph paragraph = new Paragraph();
                Run run = new Run();
                run.Append(new DocumentFormat.OpenXml.Wordprocessing.Text(plainTextContent)); 
                paragraph.Append(run);
                body.Append(paragraph);
            }

            Console.WriteLine("HTML to Word conversion completed.");
        }
    }
}
