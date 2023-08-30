using System;
using System.IO;
using System.Reflection.Metadata;
using System.Security.AccessControl;
using System.Text;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Saving;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace HTMLToWordConversionWithAspose
{
    class Program
    {
        static void Main(string[] args)
        {
            string htmlFilePath = "D:\\github\\Converter\\HTMLToWord\\HTMLToWord\\HTMLDocument.html";
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
            string wordFilePath = Path.Combine(outputDirectory, $"output_{timestamp}chartstyling.docx");

            // Converting HTML content to Word document 
            //ConvertHtmlToWordWithAspose(htmlContent, wordFilePath);



            // CONVERTING USING HTMLAGILITYPACK

            using (WordprocessingDocument doc = WordprocessingDocument.Create(wordFilePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Body());

                // Load HTML content using HtmlAgilityPack
                var htmlDocument = new HtmlDocument();
                htmlDocument.LoadHtml(htmlContent);

                // Convert parsed HTML to styled Word elements
                ConvertHtmlToWord(htmlDocument.DocumentNode.ChildNodes, body);
            }
            Console.WriteLine("HTML to Word conversion completed using Aspose.Words.");
        }

        //this method is using aspose
        static void ConvertHtmlToWordWithAspose(string htmlContent, string outputFilePath)
        {
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertHtml(htmlContent);
            //builder.InsertChart(htmlConten)

            doc.Save(outputFilePath, SaveFormat.Docx);
        }


        // this method was created for trying data binding ignore this for now
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

        //onverting html to word
        private static void ConvertHtmlToWord(HtmlNodeCollection nodes, OpenXmlElement parent)
        {
            foreach (HtmlNode node in nodes)
            {
                switch (node.NodeType)
                {
                    case HtmlNodeType.Element:
                        ApplyStylingAndConvertHtmlElementToWord(node, parent);
                        break;

                    case HtmlNodeType.Text:
                        parent.AppendChild(new Run(new Text(node.InnerText)));
                        break;
                }
            }
        }

            //trying manual styling
            private static void ApplyStylingAndConvertHtmlElementToWord(HtmlNode htmlNode, OpenXmlElement parent)
            {
                switch (htmlNode.Name.ToLower())
                {
                    case "h1":
                        parent.AppendChild(new Paragraph(
                            new ParagraphProperties(
                                new ParagraphStyleId() { Val = "Heading1" }
                            ),
                            new Run(new Text(htmlNode.InnerText))
                        ));
                        break;

                    // Handle other HTML elements and styles accordingly
                    // ...

                    default:
                        // For unsupported elements, append the inner text as plain text
                        parent.AppendChild(new Run(new Text(htmlNode.InnerText)));
                        break;
                }
            }
        

    }
}
