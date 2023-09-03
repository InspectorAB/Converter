﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Web;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace HTMLtoWord
{
    class Program
    {
        static void Main(string[] args)
        {
            string _fileCSS = "D:\\github\\Converter\\HTMLToWord2.0\\HTMLToWord2.0\\style.css";
            string _strCSS = File.ReadAllText(_fileCSS);
            string _baseURL = "http://localhost:1385/";
            string _filename = System.Guid.NewGuid().ToString() + ".doc";
            string htmlRaw = @"<table class='tbl'><thead><tr><th class='style0' colspan='2'> <img src='" + _baseURL + "img/logo.png' style='width: 180px;' /></th><th class='style1' colspan='4'><p style='font-size: 24px; padding-bottom: 2px; padding-top: 2px; font-weight: bold; margin-bottom: 1px;'>INVOICE</p> ID-2021-0024<br> Issue Date:21/09/2021<br> Delivery Date: 22/09/2021<br> Due Date:30/09/2021<br> <br><p style='font-size: 24px; padding-bottom: 2px; padding-top: 2px; font-weight: bold; margin-bottom: 1px;'>CLIENT DETAILS</p> Client 1<br> GST Number:XXXXXXXXXX</th></tr></thead><tbody><tr><td class='headstyle0' colspan='5' style='padding-top: 60px;'></td></tr><tr><td class='style3a'>ITEM</td><td class='style3a'>DESCRIPTION</td><td class='style3a'>QUANTITY</td><td class='style3a'>UNIT PRICE</td><td class='style3a'>TOTAL</td></tr><tr><td class='style3'>Item-1</td><td class='style3'>Description -1</td><td class='style3'>2 Pkt</td><td class='style3'>90.00</td><td class='style3b'>180.00</td></tr><tr><td class='style3'>Item-2</td><td class='style3'>Description-2</td><td class='style3'>5 Pkt</td><td class='style3'>35.00</td><td class='style3b'>175.00</td></tr><tr><td class='style3'>Item-3</td><td class='style3'>Description-3</td><td class='style3'>5 Kg</td><td class='style3'>50.00</td><td class='style3b'>250.00</td></tr><tr><td class='style3'>Item-4</td><td class='style3'>Description-4</td><td class='style3'>5 Kg</td><td class='style3'>150.00</td><td class='style3b'>750.00</td></tr><tr><td class='style3'>Item-5</td><td class='style3'>Description-5</td><td class='style3'>5 Kg</td><td class='style3'>100.00</td><td class='style3b'>500.00</td></tr><tr><td class='style0' colspan='2' rowspan='3'></td><td class='style3' colspan='2'>Total</td><td class='style3b'>1855.00</td></tr><tr><td class='style3' colspan='2'>GST@18%</td><td class='style3b'>333.90</td></tr><tr><td class='style3' colspan='2'>Net Payable Amount</td><td class='style3b'>2188.90</td></tr><tr><td class='style0' colspan='5' style='padding-top: 100px;'></td></tr><tr><td class='style0' colspan='5' style='background-color: aliceblue; border-radius: 2px;'><i>Note:Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</i></td></tr><tr><td class='style1' colspan='5' style='padding-top: 150px;'> Thank You<br> <b>CodeSample</b></td></tr></tbody></table>";

            StringBuilder strHTML = new StringBuilder();
            strHTML.Append("<html " +
                " xmlns:o='urn:schemas-microsoft-com:office:office'" +
                " xmlns:w='urn:schemas-microsoft-com:office:word'" +
                " xmlns='http://www.w3.org/TR/REC-html40'>" +
                "<head><title>Invoice Sample</title>");

            strHTML.Append("<xml><w:WordDocument>" +
                " <w:View>Print</w:View>" +
                " <w:Zoom>100</w:Zoom>" +
                " <w:DoNotOptimizeForBrowser/>" +
                " </w:WordDocument>" +
                " </xml>");

            strHTML.Append("<style>" + _strCSS + "</style></head>");
            strHTML.Append("<body><div class='page-settings'>" + htmlRaw + "</div></body></html>");

            var check = strHTML.ToString();

            string outputFilePath = Path.Combine("D:\\github\\Converter\\HTMLToWord2.0\\HTMLToWord2.0", _filename);
            using (WordprocessingDocument doc = WordprocessingDocument.Create(outputFilePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Create a paragraph with the HTML content
                Paragraph paragraph = new Paragraph();
                Run run = new Run(new Text(check));
                paragraph.Append(run);
                body.Append(paragraph);
            }

            Console.WriteLine("Word document created: " + outputFilePath);
            //Response.AppendHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
            //Response.AppendHeader("Content-disposition", "attachment;filename=" + _filename + "");
            //Response.Write(strHTML.ToString());
        }
    }
}