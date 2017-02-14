using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenOfficeWpfApp
{
    public class ChangeTestOfficeFileClass
    {
        private WordprocessingDocument document;

        public void ChangePackage(string filePath)
        {
            using (document = WordprocessingDocument.Open(filePath, true))
            {
                ChangeMainDocumentPart(document.MainDocumentPart);
            }
        }

        private void ChangeMainDocumentPart(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = mainDocumentPart1.Document;
            Body body = document1.GetFirstChild<Body>();

            ChangeBlueToGreen(body);
            UndelineRedWords(body);
            AddPersonalInfo(body);
        }

        private static void AddPersonalInfo(Body body)
        {
            Paragraph paragraph1 = body.Elements<Paragraph>().ElementAt(2);
            var paragraph2 = paragraph1.CloneNode(true);
            body.InsertAfter(paragraph2, paragraph1);

            Run run = paragraph1.GetFirstChild<Run>();
            RunProperties runProperties = run.GetFirstChild<RunProperties>();
            Languages usLanguage = new Languages() {Val = "en-US"};
            runProperties.Append(usLanguage);

            Text text = new Text();
            text.Text = "Andriy Zymenko, Dnipro";
            run.Append(text);
        }

        private static void ChangeBlueToGreen(Body body)
        {
            var runs = body.Elements<Paragraph>()
                .SelectMany(p => p.Elements<Run>())
                .Where(r => r.RunProperties?.Color?.Val?.Value == "0000FF");

            foreach (var run in runs)
            {
                run.RunProperties.Color.Val = "00FF00";
            }
        }

        private static void UndelineRedWords(Body body)
        {
            var text = body.InnerText;
            string[] words = Regex.Split(text, @"\s+", RegexOptions.Singleline);
            var runs = body.Elements<Paragraph>()
                .SelectMany(p => p.Elements<Run>())
                .Where(r => r.RunProperties?.Color?.Val?.Value == "FF3333" &&
                            words.Contains(Regex.Replace(r.InnerText, @"\s", ""))
                );
            foreach (var run in runs)
            {
                Underline underline = new Underline() {Val = UnderlineValues.Single};
                run.RunProperties.Append(underline);
            }
        }
    }
}
