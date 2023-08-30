using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;

namespace HTMLToWordConversionWithAspose
{
    class Program
    {
        static void Main(string[] args)
        {
            string htmlFilePath = "D:\\github\\Converter\\HTMLToWord\\HTMLToWord\\HTMLDocumentWithStyle.html";
            string dynamicHTMLFilePath = "D:\\github\\Converter\\HTMLToWord\\HTMLToWord\\Dynamic.html";
            string outputDirectory = "D:\\github\\Converter\\HTMLToWord\\HTMLToWord";



            //dynamic data conversion and data binding
            //string dTitle = "Dynamic Title";
            //string dHTMLFilePath = File.ReadAllText(dynamicHTMLFilePath);
            //byte[] htmlBytes = Encoding.UTF8.GetBytes(dHTMLFilePath);
            //string dContent = "This is dynamic content.";
            //byte[] modifiedDocumentBytes = ReplacePlaceholders(htmlBytes, dTitle, dContent);
            //File.WriteAllBytes(outputDirectory,modifiedDocumentBytes);


            // Read HTML content
            string htmlContent = File.ReadAllText(htmlFilePath);


            // File generated based on TimeStamp
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string wordFilePath = Path.Combine(outputDirectory, $"output_{timestamp}Stylingapplied.docx");

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

        static byte[] ReplacePlaceholders(byte[] templateBytes, string title, string content)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(templateBytes, 0, templateBytes.Length);

                using (WordprocessingDocument document = WordprocessingDocument.Open(stream, true))
                {
                    foreach (Text text in document.MainDocumentPart.Document.Body.Descendants<Text>())
                    {
                        if (text.Text.Contains("<<Title>>"))
                        {
                            text.Text = text.Text.Replace("<<Title>>", title);
                        }

                        if (text.Text.Contains("<<Content>>"))
                        {
                            text.Text = text.Text.Replace("<<Content>>", content);
                        }
                    }
                }

                return stream.ToArray();
            }
        }
    }
}
