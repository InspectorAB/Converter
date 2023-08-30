using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace HTMLToWordConversionWithAspose
{
    class Program
    {
        static void Main(string[] args)
        {
            string htmlFilePath = "D:\\github\\Converter\\HTMLToWord\\HTMLToWord\\HTMLDocumentWithStyleAndGraph.html";
            string outputDirectory = "D:\\github\\Converter\\HTMLToWord\\HTMLToWord"; 

            // Read HTML content
            string htmlContent = File.ReadAllText(htmlFilePath);

            // File generated based on TimeStamp
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string wordFilePath = Path.Combine(outputDirectory, $"output_{timestamp}.docx");

            // Converting HTML content to Word document 
            ConvertHtmlToWordWithAspose(htmlContent, wordFilePath);

            Console.WriteLine("HTML to Word conversion completed using Aspose.Words.");
        }

        static void ConvertHtmlToWordWithAspose(string htmlContent, string outputFilePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertHtml(htmlContent);
            

            doc.Save(outputFilePath, SaveFormat.Docx);
        }
    }
}
