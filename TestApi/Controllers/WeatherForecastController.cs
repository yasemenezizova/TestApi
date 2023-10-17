using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Interop.Word;
using Spire.Doc;
using Document = Microsoft.Office.Interop.Word.Document;
using Range = Microsoft.Office.Interop.Word.Range;
using System;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Paragraph = Spire.Doc.Documents.Paragraph;
using Section = Spire.Doc.Section;
using System.Linq;
using SautinSoft.Document.Drawing;
using Table = Spire.Doc.Table;
using System.Text.RegularExpressions;

namespace TestApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {

        //[HttpGet(Name = "GetWeatherForecast")]
        //public IActionResult Get()
        //{
        //    Aspose.Words.Document doc = new Aspose.Words.Document("Templates/qt-2.docx");
        //    DocumentBuilder builder = new DocumentBuilder(doc);
        //   // var fieldNames = new[] { "OrgName" };
        //   // var fieldValues = new object[]
        //   //{
        //   //     "Yasemenin organizationu"
        //   //};
        //   // doc.MailMerge.Execute(fieldNames, fieldValues);

        //    string searchText = "Product";
        //    string htmlContent = "<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <title>Page Title</title>\r\n    <style>\r\n        body {\r\n            font-family: 'Montserrat';\r\n                  }\r\n\r\n        li {\r\n            margin-bottom: 15px;\r\n        }\r\n\r\n        img {\r\n            width: 32%;\r\n            height: 32%;\r\n        }\r\n    </style>\r\n</head>\r\n\r\n<body>\r\n    <div>\r\n        <div style=\"width: 40%; float: left;\">\r\n            <p style=\"color: rgb(12, 185, 175, 1); font-weight: bold; font-size: 24px; margin-bottom: 24px;\">\r\n                1. Konica Minolta c257i\r\n            </p>\r\n            <p style=\"font-weight: bold; font-size:17px; margin-bottom: 16px;\">\r\n                Texniki xüsusiyət\r\n            </p>\r\n            <p>\r\n            <ul style=\"list-style-type: disc;  font-size:13px; \">\r\n                <li>Yüksək məhsuldarlıq</li>\r\n                <li>Çap mexanizminin etibarlı konstruksiyası</li>\r\n                <li>Sərf materiallarının və ehtiyat hissələrinin yüksək davamlılığı</li>\r\n                <li>Üz və arxa çapın inanılmaz dəqiqliyi</li>\r\n                <li>Sərf materiallarının və ehtiyat hissələrinin yüksək davamlılığı</li>\r\n            </ul>\r\n\r\n\r\n            </p>\r\n        </div>\r\n        <div style=\"width: 60%; float: left; margin-top: 80px;  \">\r\n            <div> <img src=\"images/image 238.jpg\" alt=\"\">\r\n                <img src=\"images/image 239.jpg\" alt=\"\">\r\n                <img src=\"images/image 240.jpg\" alt=\"\">\r\n            </div>\r\n            <div>\r\n                <p style=\" font-size:11px; margin-left: 15%; \">Quraşdırılıb: AA mətbəəsi, Max Print MMC, BB\r\n                    nəşriyyatı</p>\r\n            </div>\r\n\r\n        </div>\r\n\r\n    </div>\r\n<hr style=\"color: grey;\">\r\n</body>\r\n\r\n</html>";
        //    FindReplaceOptions options = new FindReplaceOptions();
        //    options.ReplacingCallback = new ReplaceWithHtmlCallback(htmlContent);
        //    doc.Range.Replace(searchText, "", options);
        //    MemoryStream dstStream = new MemoryStream();
        //    doc.Save(dstStream, SaveFormat.Docx);
        //    return File(dstStream.ToArray(), "application/octet-stream", $"Həkim ekspertin dəvət olunması.docx");

        //}

        [HttpGet]
        [Route("SecondTry")]
        public IActionResult SecondTry()
        {
            Spire.Doc.Document doc = new Spire.Doc.Document();
            doc.LoadFromFile("Templates/qt-1.docx");
            doc.Replace("ProductName", "LALALA", true, true);

            doc.SaveToFile("Templates/FindandReplace.docx", FileFormat.Docx2013);

            return Ok();
        }

        [HttpGet]
        [Route("DeleteMark")]
        public IActionResult DeleteMark()
        {
            var wordApp = new Application();
            var doc = wordApp.Documents.Open("C:\\Users\\yasaman.10991\\Desktop\\FindandReplace.docx");
            var range = doc.Content;
            range.Find.Text = "Evaluation Warning: The document was created with Spire.Doc for .NET.";
            range.Find.Execute(FindText: Type.Missing, ReplaceWith: Type.Missing, Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
            doc.Save();
            doc.Close();
            wordApp.Quit();
            return Ok();
        }


        [HttpGet]
        [Route("ConvertHtml")]
        public IActionResult ConvertHtml()
        {
            string filePath = @"C:\Users\yasaman.10991\Desktop\qt-2.docx";

            // Replace "t5ext" with "html"
            string findText = "Product";
            string replaceHTML = "<p>This is HTML content</p><p>You can add any HTML tags here!</p>";

            Application wordApp = new Application();
            Document doc = wordApp.Documents.Open(filePath);

            FindAndReplaceHTML(doc, findText, replaceHTML);

            // Save the changes and close the document
            doc.Save();
            doc.Close();

            // Close the Word application
            wordApp.Quit();
            return Ok();
        }


        [HttpGet]
        [Route("ConvertHtml2")]
        public IActionResult ConvertHtml2()
        {
            string filePath = @"Templates/qt-1.docx";

            Spire.Doc.Document document = new Spire.Doc.Document();
            document.LoadFromFile(filePath);

            TextSelection[] selections1 = document.FindAllString("Product", true, true);
            //Here in first parameter you need pass the string which you want to replace.
            foreach (TextSelection selection1 in selections1)
            {
                TextRange range1 = selection1.GetAsOneRange();
                Paragraph paragraph = range1.OwnerParagraph;
                int index1 = paragraph.ChildObjects.IndexOf(range1);
                paragraph.AppendHTML("<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <title>Proposal</title>\r\n    <style>\r\n        body {\r\n            font-family: 'Montserrat';\r\n            width: 12cm;           \r\n        }\r\n\r\n        li {\r\n            margin-bottom: 15px;\r\n        }\r\n\r\n        img {\r\n            width: 32%;\r\n            height: 32%;\r\n        }\r\n    </style>\r\n</head>\r\n\r\n<body>\r\n    <div>\r\n        <div style=\"width: 40%; float: left;\">\r\n            <p style=\"color: rgb(12, 185, 175, 1); font-weight: bold; font-size: 24px; margin-bottom: 24px;\">\r\n                1. Konica Minolta c257i\r\n            </p>\r\n            <p style=\"font-weight: bold; font-size:17px; margin-bottom: 16px;\">\r\n                Texniki xüsusiyət\r\n            </p>\r\n            <p>\r\n            <ul style=\"list-style-type: disc;  font-size:13px; \">\r\n                <li>Yüksək məhsuldarlıq</li>\r\n                <li>Çap mexanizminin etibarlı konstruksiyası</li>\r\n                <li>Sərf materiallarının və ehtiyat hissələrinin yüksək davamlılığı</li>\r\n                <li>Üz və arxa çapın inanılmaz dəqiqliyi</li>\r\n                <li>Sərf materiallarının və ehtiyat hissələrinin yüksək davamlılığı</li>\r\n            </ul>\r\n\r\n\r\n            </p>\r\n        </div>\r\n        <div style=\"width: 60%; float: left; margin-top: 80px;  \">\r\n            <div> <img src=\"images/image 238.jpg\" alt=\"\">\r\n                <img src=\"images/image 239.jpg\" alt=\"\">\r\n                <img src=\"images/image 240.jpg\" alt=\"\">\r\n            </div>\r\n            <div>\r\n                <p style=\" font-size:11px; margin-left: 15%; \">Quraşdırılıb: AA mətbəəsi, Max Print MMC, BB\r\n                    nəşriyyatı</p>\r\n            </div>\r\n\r\n        </div>\r\n\r\n    </div>\r\n<hr style=\"color: grey;\">\r\n</body>\r\n\r\n</html>");
                range1.OwnerParagraph.ChildObjects.Remove(range1);
            }
            document.SaveToFile(filePath, FileFormat.Docx);
            return Ok();
        }

        private static void FindAndReplaceHTML(Microsoft.Office.Interop.Word.Document doc, string findText, string replaceHTML)
        {   // Find the text "t5ext"
            Find find = doc.Content.Find;
            find.Text = findText;
            find.MatchCase = false;
            find.MatchWholeWord = true;

            while (find.Execute())
            {
                // Get the range of the found text
                Range range = find.Parent;
                if (range != null)
                {
                    // Replace the content of the found range with the HTML content
                    range.InsertXML(replaceHTML);
                }

                // Move to the next occurrence
                find.Execute(FindText: Type.Missing, Forward: true);
            }
        }

        [HttpGet]
        [Route("ConvertFirstPage")]
        public IActionResult FirstPage()
        {
            string filePath = @"Templates/qt-1.docx";

            Spire.Doc.Document document = new Spire.Doc.Document();
            document.LoadFromFile(filePath);
            document.Replace("Company", "Halal-P", true, true);
            document.Replace("ProductName", "Konica Minolat", true, true);
            document.Replace("Customer", "Yasəmen Əzizova", true, true);
            document.Replace("BeginDate", "11.12.2023", true, true);
            document.Replace("EndDate", "19.12.2023", true, true);
            document.Replace("ManagerPosition", "Banklar üzrə kurator", true, true);
            document.Replace("ManageName", "Əminə Adilzadə", true, true);
            document.Replace("ManagerPhone", "+994 50 265 85 23", true, true);
            document.Replace("ManagerEmail", "amina.adilzade@halal.az", true, true);
            document.Replace("ProposalNumber", "32423423424", true, true);
            document.SaveToFile("Replace.docx", FileFormat.Docx);
            return Ok();
        }

        [HttpGet]
        [Route("ConvertSecondPage")]
        public IActionResult SecondPage()
        {
            List<string> list = new List<string>();
            list.Add("texniki xususiyyet1");
            list.Add("texniki xususiyyet2");
            list.Add("texniki xususiyyet3");
            list.Add("texniki xususiyyet4");
            list.Add("texniki xususiyyet5");
            list.Add("texniki xususiyyet6");
            list.Add("texniki xususiyyet7");
            list.Add("texniki xususiyyet8");

            string filePath = @"Templates/qt-2.docx";

            Spire.Doc.Document document = new Spire.Doc.Document();
            document.LoadFromFile(filePath);
            for (int i = 1; i < list.Count; i++)
            {
                document.Replace("TechDesc1." + i.ToString(), list[i], true, true);
            }
            //foreach (Section section in document.Sections)
            //{
            //    for (int i = section.Paragraphs.Count - 1; i >= 0; i--)
            //    {
            //        Paragraph paragraph = section.Paragraphs[i];
            //        if (paragraph.Text.Contains("TechDesc"))
            //        {
            //            section.Paragraphs.Remove(paragraph);
            //        }
            //    }
            //}



            foreach (Section section in document.Sections)
            {
                foreach (Table table in section.Tables)
                {
                    foreach (TableRow row in table.Rows)
                    {
                        foreach (TableCell cell in row.Cells)
                        {
                            for (int i = cell.Paragraphs.Count - 1; i >= 0; i--)
                            {
                                Paragraph paragraph = cell.Paragraphs[i];
                                if (paragraph.Text.Contains("TechDesc2"))
                                {
                                    section.Paragraphs.Remove(paragraph);
                                }
                            }

                            if (cell.Paragraphs.Count > 0)
                            {
                                Paragraph paragraph = cell.Paragraphs[0];
                                string paragraphText = paragraph.Text;

                                if (paragraphText.Contains("Img1"))
                                {
                                    // Create a new paragraph with the image
                                    Paragraph newParagraph = new Paragraph(paragraph.Document);
                                    DocPicture picture = new DocPicture(paragraph.Document);
                                    picture.LoadImage(@"images/image 238.jpg");
                                    picture.Width = 100;  // Set the desired width
                                    picture.Height = 100; // Set the desired height
                                    newParagraph.ChildObjects.Add(picture);

                                    // Replace the entire cell content with the new paragraph
                                    TableCell cell2 = (TableCell)paragraph.Owner;
                                    cell2.Paragraphs.Clear();
                                    cell2.Paragraphs.Add(newParagraph);
                                }
                            }
                        }
                    }
                }
            }


            document.SaveToFile("Replace.docx", FileFormat.Docx);
            return Ok();
        }
    }
}