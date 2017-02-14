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
            Languages usLanguage = new Languages() { Val = "en-US" };
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
                Underline underline = new Underline() { Val = UnderlineValues.Single };
                run.RunProperties.Append(underline);
            }
        }

        private static void AddTable(Body body)
        {
            Run run84 = new Run();
            Break break1 = new Break() { Type = BreakValues.Page };
            run84.Append(break1);
            Paragraph paragraph5 = new Paragraph();

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Normal" };
            WidowControl widowControl1 = new WidowControl();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation1 = new Indentation() { Left = "0", Right = "0", Hanging = "0" };
            Justification justification1 = new Justification() { Val = JustificationValues.Both };
            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();

            paragraphProperties3.Append(paragraphStyleId1);
            paragraphProperties3.Append(widowControl1);
            paragraphProperties3.Append(spacingBetweenLines1);
            paragraphProperties3.Append(indentation1);
            paragraphProperties3.Append(justification1);
            paragraphProperties3.Append(paragraphMarkRunProperties3);
            paragraph5.Append(paragraphProperties3);
            //paragraph5.Append(run54);
            //paragraph5.Append(run55);
            //paragraph5.Append(run56);
            //paragraph5.Append(run57);
            //paragraph5.Append(run58);
            //paragraph5.Append(run59);
            //paragraph5.Append(run60);
            //paragraph5.Append(run61);
            //paragraph5.Append(run62);
            //paragraph5.Append(run63);
            //paragraph5.Append(run64);
            //paragraph5.Append(run65);
            //paragraph5.Append(run66);
            //paragraph5.Append(run67);
            //paragraph5.Append(run68);
            //paragraph5.Append(run69);
            //paragraph5.Append(run70);
            //paragraph5.Append(run71);
            //paragraph5.Append(run72);
            //paragraph5.Append(run73);
            //paragraph5.Append(run74);
            //paragraph5.Append(run75);
            //paragraph5.Append(run76);
            //paragraph5.Append(run77);
            //paragraph5.Append(run78);
            //paragraph5.Append(run79);
            //paragraph5.Append(run80);
            //paragraph5.Append(run81);
            //paragraph5.Append(run82);
            //paragraph5.Append(run83);
            paragraph5.Append(run84);
            //body.InsertBefore(paragraph5, sectionProperties1);

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "9638", Type = TableWidthUnitValues.Dxa };
            TableJustification tableJustification1 = new TableJustification() { Val = TableRowAlignmentValues.Left };
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 55, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "55", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 54, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "55", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 55, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableJustification1);
            tableProperties1.Append(tableIndentation1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableCellMarginDefault1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "1606" };
            GridColumn gridColumn2 = new GridColumn() { Width = "1606" };
            GridColumn gridColumn3 = new GridColumn() { Width = "1607" };
            GridColumn gridColumn4 = new GridColumn() { Width = "1606" };
            GridColumn gridColumn5 = new GridColumn() { Width = "1606" };
            GridColumn gridColumn6 = new GridColumn() { Width = "1607" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);
            tableGrid1.Append(gridColumn5);
            tableGrid1.Append(gridColumn6);

            TableRow tableRow1 = new TableRow();
            TableRowProperties tableRowProperties1 = new TableRowProperties();

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "9638", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 6 };

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(topBorder2);
            tableCellBorders1.Append(leftBorder2);
            tableCellBorders1.Append(bottomBorder2);
            tableCellBorders1.Append(rightBorder2);
            tableCellBorders1.Append(insideHorizontalBorder2);
            tableCellBorders1.Append(insideVerticalBorder2);
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin1 = new TableCellMargin();
            LeftMargin leftMargin1 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin1.Append(leftMargin1);

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(gridSpan1);
            tableCellProperties1.Append(tableCellBorders1);
            tableCellProperties1.Append(shading1);
            tableCellProperties1.Append(tableCellMargin1);

            Paragraph paragraph6 = new Paragraph();

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "TableHeading" };
            PageBreakBefore pageBreakBefore1 = new PageBreakBefore();
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts56 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold31 = new Bold();
            Bold bold32 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize32 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "26" };
            Languages languages3 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties4.Append(runFonts56);
            paragraphMarkRunProperties4.Append(bold31);
            paragraphMarkRunProperties4.Append(bold32);
            paragraphMarkRunProperties4.Append(boldComplexScript1);
            paragraphMarkRunProperties4.Append(fontSize32);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript32);
            paragraphMarkRunProperties4.Append(languages3);

            paragraphProperties4.Append(paragraphStyleId2);
            paragraphProperties4.Append(pageBreakBefore1);
            paragraphProperties4.Append(justification2);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run85 = new Run();

            RunProperties runProperties55 = new RunProperties();
            RunFonts runFonts57 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold33 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            FontSize fontSize33 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "26" };
            Languages languages4 = new Languages() { Val = "en-US" };

            runProperties55.Append(runFonts57);
            runProperties55.Append(bold33);
            runProperties55.Append(boldComplexScript2);
            runProperties55.Append(fontSize33);
            runProperties55.Append(fontSizeComplexScript33);
            runProperties55.Append(languages4);
            Text text33 = new Text();
            text33.Text = "Time Table";

            run85.Append(runProperties55);
            run85.Append(text33);

            paragraph6.Append(paragraphProperties4);
            paragraph6.Append(run85);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph6);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);

            TableRow tableRow2 = new TableRow();
            TableRowProperties tableRowProperties2 = new TableRowProperties();

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge1 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder3 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(leftBorder3);
            tableCellBorders2.Append(bottomBorder3);
            tableCellBorders2.Append(insideHorizontalBorder3);
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin2 = new TableCellMargin();
            LeftMargin leftMargin2 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin2.Append(leftMargin2);
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(verticalMerge1);
            tableCellProperties2.Append(tableCellBorders2);
            tableCellProperties2.Append(shading2);
            tableCellProperties2.Append(tableCellMargin2);
            tableCellProperties2.Append(tableCellVerticalAlignment1);

            Paragraph paragraph7 = new Paragraph();

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "TableContents" };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts58 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold34 = new Bold();
            Bold bold35 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            FontSize fontSize34 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "26" };
            Languages languages5 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties5.Append(runFonts58);
            paragraphMarkRunProperties5.Append(bold34);
            paragraphMarkRunProperties5.Append(bold35);
            paragraphMarkRunProperties5.Append(boldComplexScript3);
            paragraphMarkRunProperties5.Append(fontSize34);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript34);
            paragraphMarkRunProperties5.Append(languages5);

            paragraphProperties5.Append(paragraphStyleId3);
            paragraphProperties5.Append(justification3);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run86 = new Run();

            RunProperties runProperties56 = new RunProperties();
            RunFonts runFonts59 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold36 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            FontSize fontSize35 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "26" };
            Languages languages6 = new Languages() { Val = "en-US" };

            runProperties56.Append(runFonts59);
            runProperties56.Append(bold36);
            runProperties56.Append(boldComplexScript4);
            runProperties56.Append(fontSize35);
            runProperties56.Append(fontSizeComplexScript35);
            runProperties56.Append(languages6);
            Text text34 = new Text();
            text34.Text = "Hours";

            run86.Append(runProperties56);
            run86.Append(text34);

            paragraph7.Append(paragraphProperties5);
            paragraph7.Append(run86);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph7);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder4 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(leftBorder4);
            tableCellBorders3.Append(bottomBorder4);
            tableCellBorders3.Append(insideHorizontalBorder4);
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin3 = new TableCellMargin();
            LeftMargin leftMargin3 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin3.Append(leftMargin3);

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellBorders3);
            tableCellProperties3.Append(shading3);
            tableCellProperties3.Append(tableCellMargin3);

            Paragraph paragraph8 = new Paragraph();

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "TableContents" };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts60 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold37 = new Bold();
            Bold bold38 = new Bold();
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            FontSize fontSize36 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "26" };
            Languages languages7 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties6.Append(runFonts60);
            paragraphMarkRunProperties6.Append(bold37);
            paragraphMarkRunProperties6.Append(bold38);
            paragraphMarkRunProperties6.Append(boldComplexScript5);
            paragraphMarkRunProperties6.Append(fontSize36);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript36);
            paragraphMarkRunProperties6.Append(languages7);

            paragraphProperties6.Append(paragraphStyleId4);
            paragraphProperties6.Append(justification4);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run87 = new Run();

            RunProperties runProperties57 = new RunProperties();
            RunFonts runFonts61 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold39 = new Bold();
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            FontSize fontSize37 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "26" };
            Languages languages8 = new Languages() { Val = "en-US" };

            runProperties57.Append(runFonts61);
            runProperties57.Append(bold39);
            runProperties57.Append(boldComplexScript6);
            runProperties57.Append(fontSize37);
            runProperties57.Append(fontSizeComplexScript37);
            runProperties57.Append(languages8);
            Text text35 = new Text();
            text35.Text = "Mon";

            run87.Append(runProperties57);
            run87.Append(text35);

            paragraph8.Append(paragraphProperties6);
            paragraph8.Append(run87);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph8);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "1607", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder5 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(leftBorder5);
            tableCellBorders4.Append(bottomBorder5);
            tableCellBorders4.Append(insideHorizontalBorder5);
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin4 = new TableCellMargin();
            LeftMargin leftMargin4 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin4.Append(leftMargin4);

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(tableCellBorders4);
            tableCellProperties4.Append(shading4);
            tableCellProperties4.Append(tableCellMargin4);

            Paragraph paragraph9 = new Paragraph();

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "TableContents" };
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts62 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold40 = new Bold();
            Bold bold41 = new Bold();
            BoldComplexScript boldComplexScript7 = new BoldComplexScript();
            FontSize fontSize38 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "26" };
            Languages languages9 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties7.Append(runFonts62);
            paragraphMarkRunProperties7.Append(bold40);
            paragraphMarkRunProperties7.Append(bold41);
            paragraphMarkRunProperties7.Append(boldComplexScript7);
            paragraphMarkRunProperties7.Append(fontSize38);
            paragraphMarkRunProperties7.Append(fontSizeComplexScript38);
            paragraphMarkRunProperties7.Append(languages9);

            paragraphProperties7.Append(paragraphStyleId5);
            paragraphProperties7.Append(justification5);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            Run run88 = new Run();

            RunProperties runProperties58 = new RunProperties();
            RunFonts runFonts63 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold42 = new Bold();
            BoldComplexScript boldComplexScript8 = new BoldComplexScript();
            FontSize fontSize39 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "26" };
            Languages languages10 = new Languages() { Val = "en-US" };

            runProperties58.Append(runFonts63);
            runProperties58.Append(bold42);
            runProperties58.Append(boldComplexScript8);
            runProperties58.Append(fontSize39);
            runProperties58.Append(fontSizeComplexScript39);
            runProperties58.Append(languages10);
            Text text36 = new Text();
            text36.Text = "Tue";

            run88.Append(runProperties58);
            run88.Append(text36);

            paragraph9.Append(paragraphProperties7);
            paragraph9.Append(run88);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph9);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders5 = new TableCellBorders();
            LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder6 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders5.Append(leftBorder6);
            tableCellBorders5.Append(bottomBorder6);
            tableCellBorders5.Append(insideHorizontalBorder6);
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin5 = new TableCellMargin();
            LeftMargin leftMargin5 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin5.Append(leftMargin5);

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(tableCellBorders5);
            tableCellProperties5.Append(shading5);
            tableCellProperties5.Append(tableCellMargin5);

            Paragraph paragraph10 = new Paragraph();

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId6 = new ParagraphStyleId() { Val = "TableContents" };
            Justification justification6 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            RunFonts runFonts64 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold43 = new Bold();
            Bold bold44 = new Bold();
            BoldComplexScript boldComplexScript9 = new BoldComplexScript();
            FontSize fontSize40 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "26" };
            Languages languages11 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties8.Append(runFonts64);
            paragraphMarkRunProperties8.Append(bold43);
            paragraphMarkRunProperties8.Append(bold44);
            paragraphMarkRunProperties8.Append(boldComplexScript9);
            paragraphMarkRunProperties8.Append(fontSize40);
            paragraphMarkRunProperties8.Append(fontSizeComplexScript40);
            paragraphMarkRunProperties8.Append(languages11);

            paragraphProperties8.Append(paragraphStyleId6);
            paragraphProperties8.Append(justification6);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            Run run89 = new Run();

            RunProperties runProperties59 = new RunProperties();
            RunFonts runFonts65 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold45 = new Bold();
            BoldComplexScript boldComplexScript10 = new BoldComplexScript();
            FontSize fontSize41 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "26" };
            Languages languages12 = new Languages() { Val = "en-US" };

            runProperties59.Append(runFonts65);
            runProperties59.Append(bold45);
            runProperties59.Append(boldComplexScript10);
            runProperties59.Append(fontSize41);
            runProperties59.Append(fontSizeComplexScript41);
            runProperties59.Append(languages12);
            Text text37 = new Text();
            text37.Text = "Wed";

            run89.Append(runProperties59);
            run89.Append(text37);

            paragraph10.Append(paragraphProperties8);
            paragraph10.Append(run89);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph10);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders6 = new TableCellBorders();
            LeftBorder leftBorder7 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder7 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder7 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders6.Append(leftBorder7);
            tableCellBorders6.Append(bottomBorder7);
            tableCellBorders6.Append(insideHorizontalBorder7);
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin6 = new TableCellMargin();
            LeftMargin leftMargin6 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin6.Append(leftMargin6);

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(tableCellBorders6);
            tableCellProperties6.Append(shading6);
            tableCellProperties6.Append(tableCellMargin6);

            Paragraph paragraph11 = new Paragraph();

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId7 = new ParagraphStyleId() { Val = "TableContents" };
            Justification justification7 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts66 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold46 = new Bold();
            Bold bold47 = new Bold();
            BoldComplexScript boldComplexScript11 = new BoldComplexScript();
            FontSize fontSize42 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "26" };
            Languages languages13 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties9.Append(runFonts66);
            paragraphMarkRunProperties9.Append(bold46);
            paragraphMarkRunProperties9.Append(bold47);
            paragraphMarkRunProperties9.Append(boldComplexScript11);
            paragraphMarkRunProperties9.Append(fontSize42);
            paragraphMarkRunProperties9.Append(fontSizeComplexScript42);
            paragraphMarkRunProperties9.Append(languages13);

            paragraphProperties9.Append(paragraphStyleId7);
            paragraphProperties9.Append(justification7);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            Run run90 = new Run();

            RunProperties runProperties60 = new RunProperties();
            RunFonts runFonts67 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold48 = new Bold();
            BoldComplexScript boldComplexScript12 = new BoldComplexScript();
            FontSize fontSize43 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "26" };
            Languages languages14 = new Languages() { Val = "en-US" };

            runProperties60.Append(runFonts67);
            runProperties60.Append(bold48);
            runProperties60.Append(boldComplexScript12);
            runProperties60.Append(fontSize43);
            runProperties60.Append(fontSizeComplexScript43);
            runProperties60.Append(languages14);
            Text text38 = new Text();
            text38.Text = "Thu";

            run90.Append(runProperties60);
            run90.Append(text38);

            paragraph11.Append(paragraphProperties9);
            paragraph11.Append(run90);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph11);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "1607", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders7 = new TableCellBorders();
            LeftBorder leftBorder8 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder8 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder8 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder3 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders7.Append(leftBorder8);
            tableCellBorders7.Append(bottomBorder8);
            tableCellBorders7.Append(rightBorder3);
            tableCellBorders7.Append(insideHorizontalBorder8);
            tableCellBorders7.Append(insideVerticalBorder3);
            Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin7 = new TableCellMargin();
            LeftMargin leftMargin7 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin7.Append(leftMargin7);

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(tableCellBorders7);
            tableCellProperties7.Append(shading7);
            tableCellProperties7.Append(tableCellMargin7);

            Paragraph paragraph12 = new Paragraph();

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId8 = new ParagraphStyleId() { Val = "TableContents" };
            Justification justification8 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts68 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold49 = new Bold();
            Bold bold50 = new Bold();
            BoldComplexScript boldComplexScript13 = new BoldComplexScript();
            FontSize fontSize44 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "26" };
            Languages languages15 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties10.Append(runFonts68);
            paragraphMarkRunProperties10.Append(bold49);
            paragraphMarkRunProperties10.Append(bold50);
            paragraphMarkRunProperties10.Append(boldComplexScript13);
            paragraphMarkRunProperties10.Append(fontSize44);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript44);
            paragraphMarkRunProperties10.Append(languages15);

            paragraphProperties10.Append(paragraphStyleId8);
            paragraphProperties10.Append(justification8);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            Run run91 = new Run();

            RunProperties runProperties61 = new RunProperties();
            RunFonts runFonts69 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold51 = new Bold();
            BoldComplexScript boldComplexScript14 = new BoldComplexScript();
            FontSize fontSize45 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "26" };
            Languages languages16 = new Languages() { Val = "en-US" };

            runProperties61.Append(runFonts69);
            runProperties61.Append(bold51);
            runProperties61.Append(boldComplexScript14);
            runProperties61.Append(fontSize45);
            runProperties61.Append(fontSizeComplexScript45);
            runProperties61.Append(languages16);
            Text text39 = new Text();
            text39.Text = "Fri";

            run91.Append(runProperties61);
            run91.Append(text39);

            paragraph12.Append(paragraphProperties10);
            paragraph12.Append(run91);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph12);

            tableRow2.Append(tableRowProperties2);
            tableRow2.Append(tableCell2);
            tableRow2.Append(tableCell3);
            tableRow2.Append(tableCell4);
            tableRow2.Append(tableCell5);
            tableRow2.Append(tableCell6);
            tableRow2.Append(tableCell7);

            TableRow tableRow3 = new TableRow();
            TableRowProperties tableRowProperties3 = new TableRowProperties();

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge2 = new VerticalMerge() { Val = MergedCellValues.Continue };

            TableCellBorders tableCellBorders8 = new TableCellBorders();
            LeftBorder leftBorder9 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder9 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder9 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders8.Append(leftBorder9);
            tableCellBorders8.Append(bottomBorder9);
            tableCellBorders8.Append(insideHorizontalBorder9);
            Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin8 = new TableCellMargin();
            LeftMargin leftMargin8 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin8.Append(leftMargin8);

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(verticalMerge2);
            tableCellProperties8.Append(tableCellBorders8);
            tableCellProperties8.Append(shading8);
            tableCellProperties8.Append(tableCellMargin8);

            Paragraph paragraph13 = new Paragraph();

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId9 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts70 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize46 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties11.Append(runFonts70);
            paragraphMarkRunProperties11.Append(fontSize46);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript46);

            paragraphProperties11.Append(paragraphStyleId9);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run92 = new Run();

            RunProperties runProperties62 = new RunProperties();
            RunFonts runFonts71 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize47 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "26" };

            runProperties62.Append(runFonts71);
            runProperties62.Append(fontSize47);
            runProperties62.Append(fontSizeComplexScript47);

            run92.Append(runProperties62);

            paragraph13.Append(paragraphProperties11);
            paragraph13.Append(run92);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph13);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders9 = new TableCellBorders();
            LeftBorder leftBorder10 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder10 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder10 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders9.Append(leftBorder10);
            tableCellBorders9.Append(bottomBorder10);
            tableCellBorders9.Append(insideHorizontalBorder10);
            Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin9 = new TableCellMargin();
            LeftMargin leftMargin9 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin9.Append(leftMargin9);

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(tableCellBorders9);
            tableCellProperties9.Append(shading9);
            tableCellProperties9.Append(tableCellMargin9);

            Paragraph paragraph14 = new Paragraph();

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId10 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts72 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize48 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "26" };
            Languages languages17 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties12.Append(runFonts72);
            paragraphMarkRunProperties12.Append(fontSize48);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript48);
            paragraphMarkRunProperties12.Append(languages17);

            paragraphProperties12.Append(paragraphStyleId10);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run93 = new Run();

            RunProperties runProperties63 = new RunProperties();
            RunFonts runFonts73 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize49 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "26" };
            Languages languages18 = new Languages() { Val = "en-US" };

            runProperties63.Append(runFonts73);
            runProperties63.Append(fontSize49);
            runProperties63.Append(fontSizeComplexScript49);
            runProperties63.Append(languages18);
            Text text40 = new Text();
            text40.Text = "Science";

            run93.Append(runProperties63);
            run93.Append(text40);

            paragraph14.Append(paragraphProperties12);
            paragraph14.Append(run93);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph14);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "1607", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders10 = new TableCellBorders();
            LeftBorder leftBorder11 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder11 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder11 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders10.Append(leftBorder11);
            tableCellBorders10.Append(bottomBorder11);
            tableCellBorders10.Append(insideHorizontalBorder11);
            Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin10 = new TableCellMargin();
            LeftMargin leftMargin10 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin10.Append(leftMargin10);

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(tableCellBorders10);
            tableCellProperties10.Append(shading10);
            tableCellProperties10.Append(tableCellMargin10);

            Paragraph paragraph15 = new Paragraph();

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId11 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunFonts runFonts74 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize50 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "26" };
            Languages languages19 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties13.Append(runFonts74);
            paragraphMarkRunProperties13.Append(fontSize50);
            paragraphMarkRunProperties13.Append(fontSizeComplexScript50);
            paragraphMarkRunProperties13.Append(languages19);

            paragraphProperties13.Append(paragraphStyleId11);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

            Run run94 = new Run();

            RunProperties runProperties64 = new RunProperties();
            RunFonts runFonts75 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize51 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "26" };
            Languages languages20 = new Languages() { Val = "en-US" };

            runProperties64.Append(runFonts75);
            runProperties64.Append(fontSize51);
            runProperties64.Append(fontSizeComplexScript51);
            runProperties64.Append(languages20);
            Text text41 = new Text();
            text41.Text = "Maths";

            run94.Append(runProperties64);
            run94.Append(text41);

            paragraph15.Append(paragraphProperties13);
            paragraph15.Append(run94);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph15);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders11 = new TableCellBorders();
            LeftBorder leftBorder12 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder12 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder12 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders11.Append(leftBorder12);
            tableCellBorders11.Append(bottomBorder12);
            tableCellBorders11.Append(insideHorizontalBorder12);
            Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin11 = new TableCellMargin();
            LeftMargin leftMargin11 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin11.Append(leftMargin11);

            tableCellProperties11.Append(tableCellWidth11);
            tableCellProperties11.Append(tableCellBorders11);
            tableCellProperties11.Append(shading11);
            tableCellProperties11.Append(tableCellMargin11);

            Paragraph paragraph16 = new Paragraph();

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId12 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts76 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize52 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "26" };
            Languages languages21 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties14.Append(runFonts76);
            paragraphMarkRunProperties14.Append(fontSize52);
            paragraphMarkRunProperties14.Append(fontSizeComplexScript52);
            paragraphMarkRunProperties14.Append(languages21);

            paragraphProperties14.Append(paragraphStyleId12);
            paragraphProperties14.Append(paragraphMarkRunProperties14);

            Run run95 = new Run();

            RunProperties runProperties65 = new RunProperties();
            RunFonts runFonts77 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize53 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "26" };
            Languages languages22 = new Languages() { Val = "en-US" };

            runProperties65.Append(runFonts77);
            runProperties65.Append(fontSize53);
            runProperties65.Append(fontSizeComplexScript53);
            runProperties65.Append(languages22);
            Text text42 = new Text();
            text42.Text = "Science";

            run95.Append(runProperties65);
            run95.Append(text42);

            paragraph16.Append(paragraphProperties14);
            paragraph16.Append(run95);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph16);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders12 = new TableCellBorders();
            LeftBorder leftBorder13 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder13 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder13 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders12.Append(leftBorder13);
            tableCellBorders12.Append(bottomBorder13);
            tableCellBorders12.Append(insideHorizontalBorder13);
            Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin12 = new TableCellMargin();
            LeftMargin leftMargin12 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin12.Append(leftMargin12);

            tableCellProperties12.Append(tableCellWidth12);
            tableCellProperties12.Append(tableCellBorders12);
            tableCellProperties12.Append(shading12);
            tableCellProperties12.Append(tableCellMargin12);

            Paragraph paragraph17 = new Paragraph();

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId13 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            RunFonts runFonts78 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize54 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "26" };
            Languages languages23 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties15.Append(runFonts78);
            paragraphMarkRunProperties15.Append(fontSize54);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript54);
            paragraphMarkRunProperties15.Append(languages23);

            paragraphProperties15.Append(paragraphStyleId13);
            paragraphProperties15.Append(paragraphMarkRunProperties15);

            Run run96 = new Run();

            RunProperties runProperties66 = new RunProperties();
            RunFonts runFonts79 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize55 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "26" };
            Languages languages24 = new Languages() { Val = "en-US" };

            runProperties66.Append(runFonts79);
            runProperties66.Append(fontSize55);
            runProperties66.Append(fontSizeComplexScript55);
            runProperties66.Append(languages24);
            Text text43 = new Text();
            text43.Text = "Maths";

            run96.Append(runProperties66);
            run96.Append(text43);

            paragraph17.Append(paragraphProperties15);
            paragraph17.Append(run96);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph17);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "1607", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders13 = new TableCellBorders();
            LeftBorder leftBorder14 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder14 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder14 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder4 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders13.Append(leftBorder14);
            tableCellBorders13.Append(bottomBorder14);
            tableCellBorders13.Append(rightBorder4);
            tableCellBorders13.Append(insideHorizontalBorder14);
            tableCellBorders13.Append(insideVerticalBorder4);
            Shading shading13 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin13 = new TableCellMargin();
            LeftMargin leftMargin13 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin13.Append(leftMargin13);

            tableCellProperties13.Append(tableCellWidth13);
            tableCellProperties13.Append(tableCellBorders13);
            tableCellProperties13.Append(shading13);
            tableCellProperties13.Append(tableCellMargin13);

            Paragraph paragraph18 = new Paragraph();

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId14 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunFonts runFonts80 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize56 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "26" };
            Languages languages25 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties16.Append(runFonts80);
            paragraphMarkRunProperties16.Append(fontSize56);
            paragraphMarkRunProperties16.Append(fontSizeComplexScript56);
            paragraphMarkRunProperties16.Append(languages25);

            paragraphProperties16.Append(paragraphStyleId14);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            Run run97 = new Run();

            RunProperties runProperties67 = new RunProperties();
            RunFonts runFonts81 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize57 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "26" };
            Languages languages26 = new Languages() { Val = "en-US" };

            runProperties67.Append(runFonts81);
            runProperties67.Append(fontSize57);
            runProperties67.Append(fontSizeComplexScript57);
            runProperties67.Append(languages26);
            Text text44 = new Text();
            text44.Text = "Arts";

            run97.Append(runProperties67);
            run97.Append(text44);

            paragraph18.Append(paragraphProperties16);
            paragraph18.Append(run97);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph18);

            tableRow3.Append(tableRowProperties3);
            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);
            tableRow3.Append(tableCell10);
            tableRow3.Append(tableCell11);
            tableRow3.Append(tableCell12);
            tableRow3.Append(tableCell13);

            TableRow tableRow4 = new TableRow();
            TableRowProperties tableRowProperties4 = new TableRowProperties();

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge3 = new VerticalMerge() { Val = MergedCellValues.Continue };

            TableCellBorders tableCellBorders14 = new TableCellBorders();
            LeftBorder leftBorder15 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder15 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder15 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders14.Append(leftBorder15);
            tableCellBorders14.Append(bottomBorder15);
            tableCellBorders14.Append(insideHorizontalBorder15);
            Shading shading14 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin14 = new TableCellMargin();
            LeftMargin leftMargin14 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin14.Append(leftMargin14);

            tableCellProperties14.Append(tableCellWidth14);
            tableCellProperties14.Append(verticalMerge3);
            tableCellProperties14.Append(tableCellBorders14);
            tableCellProperties14.Append(shading14);
            tableCellProperties14.Append(tableCellMargin14);

            Paragraph paragraph19 = new Paragraph();

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId15 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            RunFonts runFonts82 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize58 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties17.Append(runFonts82);
            paragraphMarkRunProperties17.Append(fontSize58);
            paragraphMarkRunProperties17.Append(fontSizeComplexScript58);

            paragraphProperties17.Append(paragraphStyleId15);
            paragraphProperties17.Append(paragraphMarkRunProperties17);

            Run run98 = new Run();

            RunProperties runProperties68 = new RunProperties();
            RunFonts runFonts83 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize59 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript59 = new FontSizeComplexScript() { Val = "26" };

            runProperties68.Append(runFonts83);
            runProperties68.Append(fontSize59);
            runProperties68.Append(fontSizeComplexScript59);

            run98.Append(runProperties68);

            paragraph19.Append(paragraphProperties17);
            paragraph19.Append(run98);

            tableCell14.Append(tableCellProperties14);
            tableCell14.Append(paragraph19);

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders15 = new TableCellBorders();
            LeftBorder leftBorder16 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder16 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder16 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders15.Append(leftBorder16);
            tableCellBorders15.Append(bottomBorder16);
            tableCellBorders15.Append(insideHorizontalBorder16);
            Shading shading15 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin15 = new TableCellMargin();
            LeftMargin leftMargin15 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin15.Append(leftMargin15);

            tableCellProperties15.Append(tableCellWidth15);
            tableCellProperties15.Append(tableCellBorders15);
            tableCellProperties15.Append(shading15);
            tableCellProperties15.Append(tableCellMargin15);

            Paragraph paragraph20 = new Paragraph();

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId16 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts84 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize60 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript60 = new FontSizeComplexScript() { Val = "26" };
            Languages languages27 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties18.Append(runFonts84);
            paragraphMarkRunProperties18.Append(fontSize60);
            paragraphMarkRunProperties18.Append(fontSizeComplexScript60);
            paragraphMarkRunProperties18.Append(languages27);

            paragraphProperties18.Append(paragraphStyleId16);
            paragraphProperties18.Append(paragraphMarkRunProperties18);

            Run run99 = new Run();

            RunProperties runProperties69 = new RunProperties();
            RunFonts runFonts85 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize61 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript61 = new FontSizeComplexScript() { Val = "26" };
            Languages languages28 = new Languages() { Val = "en-US" };

            runProperties69.Append(runFonts85);
            runProperties69.Append(fontSize61);
            runProperties69.Append(fontSizeComplexScript61);
            runProperties69.Append(languages28);
            Text text45 = new Text();
            text45.Text = "Social";

            run99.Append(runProperties69);
            run99.Append(text45);

            paragraph20.Append(paragraphProperties18);
            paragraph20.Append(run99);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph20);

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "1607", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders16 = new TableCellBorders();
            LeftBorder leftBorder17 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder17 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder17 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders16.Append(leftBorder17);
            tableCellBorders16.Append(bottomBorder17);
            tableCellBorders16.Append(insideHorizontalBorder17);
            Shading shading16 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin16 = new TableCellMargin();
            LeftMargin leftMargin16 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin16.Append(leftMargin16);

            tableCellProperties16.Append(tableCellWidth16);
            tableCellProperties16.Append(tableCellBorders16);
            tableCellProperties16.Append(shading16);
            tableCellProperties16.Append(tableCellMargin16);

            Paragraph paragraph21 = new Paragraph();

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId17 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts86 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize62 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript62 = new FontSizeComplexScript() { Val = "26" };
            Languages languages29 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties19.Append(runFonts86);
            paragraphMarkRunProperties19.Append(fontSize62);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript62);
            paragraphMarkRunProperties19.Append(languages29);

            paragraphProperties19.Append(paragraphStyleId17);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run100 = new Run();

            RunProperties runProperties70 = new RunProperties();
            RunFonts runFonts87 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize63 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "26" };
            Languages languages30 = new Languages() { Val = "en-US" };

            runProperties70.Append(runFonts87);
            runProperties70.Append(fontSize63);
            runProperties70.Append(fontSizeComplexScript63);
            runProperties70.Append(languages30);
            Text text46 = new Text();
            text46.Text = "History";

            run100.Append(runProperties70);
            run100.Append(text46);

            paragraph21.Append(paragraphProperties19);
            paragraph21.Append(run100);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph21);

            TableCell tableCell17 = new TableCell();

            TableCellProperties tableCellProperties17 = new TableCellProperties();
            TableCellWidth tableCellWidth17 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders17 = new TableCellBorders();
            LeftBorder leftBorder18 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder18 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder18 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders17.Append(leftBorder18);
            tableCellBorders17.Append(bottomBorder18);
            tableCellBorders17.Append(insideHorizontalBorder18);
            Shading shading17 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin17 = new TableCellMargin();
            LeftMargin leftMargin17 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin17.Append(leftMargin17);

            tableCellProperties17.Append(tableCellWidth17);
            tableCellProperties17.Append(tableCellBorders17);
            tableCellProperties17.Append(shading17);
            tableCellProperties17.Append(tableCellMargin17);

            Paragraph paragraph22 = new Paragraph();

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId18 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts88 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize64 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "26" };
            Languages languages31 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties20.Append(runFonts88);
            paragraphMarkRunProperties20.Append(fontSize64);
            paragraphMarkRunProperties20.Append(fontSizeComplexScript64);
            paragraphMarkRunProperties20.Append(languages31);

            paragraphProperties20.Append(paragraphStyleId18);
            paragraphProperties20.Append(paragraphMarkRunProperties20);

            Run run101 = new Run();

            RunProperties runProperties71 = new RunProperties();
            RunFonts runFonts89 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize65 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript65 = new FontSizeComplexScript() { Val = "26" };
            Languages languages32 = new Languages() { Val = "en-US" };

            runProperties71.Append(runFonts89);
            runProperties71.Append(fontSize65);
            runProperties71.Append(fontSizeComplexScript65);
            runProperties71.Append(languages32);
            Text text47 = new Text();
            text47.Text = "English";

            run101.Append(runProperties71);
            run101.Append(text47);

            paragraph22.Append(paragraphProperties20);
            paragraph22.Append(run101);

            tableCell17.Append(tableCellProperties17);
            tableCell17.Append(paragraph22);

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders18 = new TableCellBorders();
            LeftBorder leftBorder19 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder19 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder19 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders18.Append(leftBorder19);
            tableCellBorders18.Append(bottomBorder19);
            tableCellBorders18.Append(insideHorizontalBorder19);
            Shading shading18 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin18 = new TableCellMargin();
            LeftMargin leftMargin18 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin18.Append(leftMargin18);

            tableCellProperties18.Append(tableCellWidth18);
            tableCellProperties18.Append(tableCellBorders18);
            tableCellProperties18.Append(shading18);
            tableCellProperties18.Append(tableCellMargin18);

            Paragraph paragraph23 = new Paragraph();

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId19 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts90 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize66 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript66 = new FontSizeComplexScript() { Val = "26" };
            Languages languages33 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties21.Append(runFonts90);
            paragraphMarkRunProperties21.Append(fontSize66);
            paragraphMarkRunProperties21.Append(fontSizeComplexScript66);
            paragraphMarkRunProperties21.Append(languages33);

            paragraphProperties21.Append(paragraphStyleId19);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run102 = new Run();

            RunProperties runProperties72 = new RunProperties();
            RunFonts runFonts91 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize67 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript67 = new FontSizeComplexScript() { Val = "26" };
            Languages languages34 = new Languages() { Val = "en-US" };

            runProperties72.Append(runFonts91);
            runProperties72.Append(fontSize67);
            runProperties72.Append(fontSizeComplexScript67);
            runProperties72.Append(languages34);
            Text text48 = new Text();
            text48.Text = "Social";

            run102.Append(runProperties72);
            run102.Append(text48);

            paragraph23.Append(paragraphProperties21);
            paragraph23.Append(run102);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph23);

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "1607", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders19 = new TableCellBorders();
            LeftBorder leftBorder20 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder20 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            RightBorder rightBorder5 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder20 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder5 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders19.Append(leftBorder20);
            tableCellBorders19.Append(bottomBorder20);
            tableCellBorders19.Append(rightBorder5);
            tableCellBorders19.Append(insideHorizontalBorder20);
            tableCellBorders19.Append(insideVerticalBorder5);
            Shading shading19 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin19 = new TableCellMargin();
            LeftMargin leftMargin19 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin19.Append(leftMargin19);

            tableCellProperties19.Append(tableCellWidth19);
            tableCellProperties19.Append(tableCellBorders19);
            tableCellProperties19.Append(shading19);
            tableCellProperties19.Append(tableCellMargin19);

            Paragraph paragraph24 = new Paragraph();

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId20 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts92 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize68 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript68 = new FontSizeComplexScript() { Val = "26" };
            Languages languages35 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties22.Append(runFonts92);
            paragraphMarkRunProperties22.Append(fontSize68);
            paragraphMarkRunProperties22.Append(fontSizeComplexScript68);
            paragraphMarkRunProperties22.Append(languages35);

            paragraphProperties22.Append(paragraphStyleId20);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            Run run103 = new Run();

            RunProperties runProperties73 = new RunProperties();
            RunFonts runFonts93 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize69 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript69 = new FontSizeComplexScript() { Val = "26" };
            Languages languages36 = new Languages() { Val = "en-US" };

            runProperties73.Append(runFonts93);
            runProperties73.Append(fontSize69);
            runProperties73.Append(fontSizeComplexScript69);
            runProperties73.Append(languages36);
            Text text49 = new Text();
            text49.Text = "Sports";

            run103.Append(runProperties73);
            run103.Append(text49);

            paragraph24.Append(paragraphProperties22);
            paragraph24.Append(run103);

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph24);

            tableRow4.Append(tableRowProperties4);
            tableRow4.Append(tableCell14);
            tableRow4.Append(tableCell15);
            tableRow4.Append(tableCell16);
            tableRow4.Append(tableCell17);
            tableRow4.Append(tableCell18);
            tableRow4.Append(tableCell19);

            TableRow tableRow5 = new TableRow();
            TableRowProperties tableRowProperties5 = new TableRowProperties();

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge4 = new VerticalMerge() { Val = MergedCellValues.Continue };

            TableCellBorders tableCellBorders20 = new TableCellBorders();
            LeftBorder leftBorder21 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder21 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder21 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders20.Append(leftBorder21);
            tableCellBorders20.Append(bottomBorder21);
            tableCellBorders20.Append(insideHorizontalBorder21);
            Shading shading20 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin20 = new TableCellMargin();
            LeftMargin leftMargin20 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin20.Append(leftMargin20);

            tableCellProperties20.Append(tableCellWidth20);
            tableCellProperties20.Append(verticalMerge4);
            tableCellProperties20.Append(tableCellBorders20);
            tableCellProperties20.Append(shading20);
            tableCellProperties20.Append(tableCellMargin20);

            Paragraph paragraph25 = new Paragraph();

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId21 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts94 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize70 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript70 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties23.Append(runFonts94);
            paragraphMarkRunProperties23.Append(fontSize70);
            paragraphMarkRunProperties23.Append(fontSizeComplexScript70);

            paragraphProperties23.Append(paragraphStyleId21);
            paragraphProperties23.Append(paragraphMarkRunProperties23);

            Run run104 = new Run();

            RunProperties runProperties74 = new RunProperties();
            RunFonts runFonts95 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize71 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript71 = new FontSizeComplexScript() { Val = "26" };

            runProperties74.Append(runFonts95);
            runProperties74.Append(fontSize71);
            runProperties74.Append(fontSizeComplexScript71);

            run104.Append(runProperties74);

            paragraph25.Append(paragraphProperties23);
            paragraph25.Append(run104);

            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph25);

            TableCell tableCell21 = new TableCell();

            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "8032", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan2 = new GridSpan() { Val = 5 };

            TableCellBorders tableCellBorders21 = new TableCellBorders();
            LeftBorder leftBorder22 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder22 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            RightBorder rightBorder6 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder22 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder6 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders21.Append(leftBorder22);
            tableCellBorders21.Append(bottomBorder22);
            tableCellBorders21.Append(rightBorder6);
            tableCellBorders21.Append(insideHorizontalBorder22);
            tableCellBorders21.Append(insideVerticalBorder6);
            Shading shading21 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin21 = new TableCellMargin();
            LeftMargin leftMargin21 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin21.Append(leftMargin21);

            tableCellProperties21.Append(tableCellWidth21);
            tableCellProperties21.Append(gridSpan2);
            tableCellProperties21.Append(tableCellBorders21);
            tableCellProperties21.Append(shading21);
            tableCellProperties21.Append(tableCellMargin21);

            Paragraph paragraph26 = new Paragraph();

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId22 = new ParagraphStyleId() { Val = "TableContents" };
            Justification justification9 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            RunFonts runFonts96 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold52 = new Bold();
            Bold bold53 = new Bold();
            BoldComplexScript boldComplexScript15 = new BoldComplexScript();
            FontSize fontSize72 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript72 = new FontSizeComplexScript() { Val = "26" };
            Languages languages37 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties24.Append(runFonts96);
            paragraphMarkRunProperties24.Append(bold52);
            paragraphMarkRunProperties24.Append(bold53);
            paragraphMarkRunProperties24.Append(boldComplexScript15);
            paragraphMarkRunProperties24.Append(fontSize72);
            paragraphMarkRunProperties24.Append(fontSizeComplexScript72);
            paragraphMarkRunProperties24.Append(languages37);

            paragraphProperties24.Append(paragraphStyleId22);
            paragraphProperties24.Append(justification9);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            Run run105 = new Run();

            RunProperties runProperties75 = new RunProperties();
            RunFonts runFonts97 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Bold bold54 = new Bold();
            BoldComplexScript boldComplexScript16 = new BoldComplexScript();
            FontSize fontSize73 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript73 = new FontSizeComplexScript() { Val = "26" };
            Languages languages38 = new Languages() { Val = "en-US" };

            runProperties75.Append(runFonts97);
            runProperties75.Append(bold54);
            runProperties75.Append(boldComplexScript16);
            runProperties75.Append(fontSize73);
            runProperties75.Append(fontSizeComplexScript73);
            runProperties75.Append(languages38);
            Text text50 = new Text();
            text50.Text = "Lunch";

            run105.Append(runProperties75);
            run105.Append(text50);

            paragraph26.Append(paragraphProperties24);
            paragraph26.Append(run105);

            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph26);

            tableRow5.Append(tableRowProperties5);
            tableRow5.Append(tableCell20);
            tableRow5.Append(tableCell21);

            TableRow tableRow6 = new TableRow();
            TableRowProperties tableRowProperties6 = new TableRowProperties();

            TableCell tableCell22 = new TableCell();

            TableCellProperties tableCellProperties22 = new TableCellProperties();
            TableCellWidth tableCellWidth22 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge5 = new VerticalMerge() { Val = MergedCellValues.Continue };

            TableCellBorders tableCellBorders22 = new TableCellBorders();
            LeftBorder leftBorder23 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder23 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder23 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders22.Append(leftBorder23);
            tableCellBorders22.Append(bottomBorder23);
            tableCellBorders22.Append(insideHorizontalBorder23);
            Shading shading22 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin22 = new TableCellMargin();
            LeftMargin leftMargin22 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin22.Append(leftMargin22);

            tableCellProperties22.Append(tableCellWidth22);
            tableCellProperties22.Append(verticalMerge5);
            tableCellProperties22.Append(tableCellBorders22);
            tableCellProperties22.Append(shading22);
            tableCellProperties22.Append(tableCellMargin22);

            Paragraph paragraph27 = new Paragraph();

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId23 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts98 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize74 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript74 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties25.Append(runFonts98);
            paragraphMarkRunProperties25.Append(fontSize74);
            paragraphMarkRunProperties25.Append(fontSizeComplexScript74);

            paragraphProperties25.Append(paragraphStyleId23);
            paragraphProperties25.Append(paragraphMarkRunProperties25);

            Run run106 = new Run();

            RunProperties runProperties76 = new RunProperties();
            RunFonts runFonts99 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize75 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript75 = new FontSizeComplexScript() { Val = "26" };

            runProperties76.Append(runFonts99);
            runProperties76.Append(fontSize75);
            runProperties76.Append(fontSizeComplexScript75);

            run106.Append(runProperties76);

            paragraph27.Append(paragraphProperties25);
            paragraph27.Append(run106);

            tableCell22.Append(tableCellProperties22);
            tableCell22.Append(paragraph27);

            TableCell tableCell23 = new TableCell();

            TableCellProperties tableCellProperties23 = new TableCellProperties();
            TableCellWidth tableCellWidth23 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders23 = new TableCellBorders();
            LeftBorder leftBorder24 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder24 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder24 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders23.Append(leftBorder24);
            tableCellBorders23.Append(bottomBorder24);
            tableCellBorders23.Append(insideHorizontalBorder24);
            Shading shading23 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin23 = new TableCellMargin();
            LeftMargin leftMargin23 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin23.Append(leftMargin23);

            tableCellProperties23.Append(tableCellWidth23);
            tableCellProperties23.Append(tableCellBorders23);
            tableCellProperties23.Append(shading23);
            tableCellProperties23.Append(tableCellMargin23);

            Paragraph paragraph28 = new Paragraph();

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId24 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            RunFonts runFonts100 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize76 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript76 = new FontSizeComplexScript() { Val = "26" };
            Languages languages39 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties26.Append(runFonts100);
            paragraphMarkRunProperties26.Append(fontSize76);
            paragraphMarkRunProperties26.Append(fontSizeComplexScript76);
            paragraphMarkRunProperties26.Append(languages39);

            paragraphProperties26.Append(paragraphStyleId24);
            paragraphProperties26.Append(paragraphMarkRunProperties26);

            Run run107 = new Run();

            RunProperties runProperties77 = new RunProperties();
            RunFonts runFonts101 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize77 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript77 = new FontSizeComplexScript() { Val = "26" };
            Languages languages40 = new Languages() { Val = "en-US" };

            runProperties77.Append(runFonts101);
            runProperties77.Append(fontSize77);
            runProperties77.Append(fontSizeComplexScript77);
            runProperties77.Append(languages40);
            Text text51 = new Text();
            text51.Text = "Science";

            run107.Append(runProperties77);
            run107.Append(text51);

            paragraph28.Append(paragraphProperties26);
            paragraph28.Append(run107);

            tableCell23.Append(tableCellProperties23);
            tableCell23.Append(paragraph28);

            TableCell tableCell24 = new TableCell();

            TableCellProperties tableCellProperties24 = new TableCellProperties();
            TableCellWidth tableCellWidth24 = new TableCellWidth() { Width = "1607", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders24 = new TableCellBorders();
            LeftBorder leftBorder25 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder25 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder25 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders24.Append(leftBorder25);
            tableCellBorders24.Append(bottomBorder25);
            tableCellBorders24.Append(insideHorizontalBorder25);
            Shading shading24 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin24 = new TableCellMargin();
            LeftMargin leftMargin24 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin24.Append(leftMargin24);

            tableCellProperties24.Append(tableCellWidth24);
            tableCellProperties24.Append(tableCellBorders24);
            tableCellProperties24.Append(shading24);
            tableCellProperties24.Append(tableCellMargin24);

            Paragraph paragraph29 = new Paragraph();

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId25 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            RunFonts runFonts102 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize78 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript78 = new FontSizeComplexScript() { Val = "26" };
            Languages languages41 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties27.Append(runFonts102);
            paragraphMarkRunProperties27.Append(fontSize78);
            paragraphMarkRunProperties27.Append(fontSizeComplexScript78);
            paragraphMarkRunProperties27.Append(languages41);

            paragraphProperties27.Append(paragraphStyleId25);
            paragraphProperties27.Append(paragraphMarkRunProperties27);

            Run run108 = new Run();

            RunProperties runProperties78 = new RunProperties();
            RunFonts runFonts103 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize79 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript79 = new FontSizeComplexScript() { Val = "26" };
            Languages languages42 = new Languages() { Val = "en-US" };

            runProperties78.Append(runFonts103);
            runProperties78.Append(fontSize79);
            runProperties78.Append(fontSizeComplexScript79);
            runProperties78.Append(languages42);
            Text text52 = new Text();
            text52.Text = "Maths";

            run108.Append(runProperties78);
            run108.Append(text52);

            paragraph29.Append(paragraphProperties27);
            paragraph29.Append(run108);

            tableCell24.Append(tableCellProperties24);
            tableCell24.Append(paragraph29);

            TableCell tableCell25 = new TableCell();

            TableCellProperties tableCellProperties25 = new TableCellProperties();
            TableCellWidth tableCellWidth25 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders25 = new TableCellBorders();
            LeftBorder leftBorder26 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder26 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder26 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders25.Append(leftBorder26);
            tableCellBorders25.Append(bottomBorder26);
            tableCellBorders25.Append(insideHorizontalBorder26);
            Shading shading25 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin25 = new TableCellMargin();
            LeftMargin leftMargin25 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin25.Append(leftMargin25);

            tableCellProperties25.Append(tableCellWidth25);
            tableCellProperties25.Append(tableCellBorders25);
            tableCellProperties25.Append(shading25);
            tableCellProperties25.Append(tableCellMargin25);

            Paragraph paragraph30 = new Paragraph();

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId26 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            RunFonts runFonts104 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize80 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript80 = new FontSizeComplexScript() { Val = "26" };
            Languages languages43 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties28.Append(runFonts104);
            paragraphMarkRunProperties28.Append(fontSize80);
            paragraphMarkRunProperties28.Append(fontSizeComplexScript80);
            paragraphMarkRunProperties28.Append(languages43);

            paragraphProperties28.Append(paragraphStyleId26);
            paragraphProperties28.Append(paragraphMarkRunProperties28);

            Run run109 = new Run();

            RunProperties runProperties79 = new RunProperties();
            RunFonts runFonts105 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize81 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript81 = new FontSizeComplexScript() { Val = "26" };
            Languages languages44 = new Languages() { Val = "en-US" };

            runProperties79.Append(runFonts105);
            runProperties79.Append(fontSize81);
            runProperties79.Append(fontSizeComplexScript81);
            runProperties79.Append(languages44);
            Text text53 = new Text();
            text53.Text = "Science";

            run109.Append(runProperties79);
            run109.Append(text53);

            paragraph30.Append(paragraphProperties28);
            paragraph30.Append(run109);

            tableCell25.Append(tableCellProperties25);
            tableCell25.Append(paragraph30);

            TableCell tableCell26 = new TableCell();

            TableCellProperties tableCellProperties26 = new TableCellProperties();
            TableCellWidth tableCellWidth26 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders26 = new TableCellBorders();
            LeftBorder leftBorder27 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder27 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder27 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders26.Append(leftBorder27);
            tableCellBorders26.Append(bottomBorder27);
            tableCellBorders26.Append(insideHorizontalBorder27);
            Shading shading26 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin26 = new TableCellMargin();
            LeftMargin leftMargin26 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin26.Append(leftMargin26);

            tableCellProperties26.Append(tableCellWidth26);
            tableCellProperties26.Append(tableCellBorders26);
            tableCellProperties26.Append(shading26);
            tableCellProperties26.Append(tableCellMargin26);

            Paragraph paragraph31 = new Paragraph();

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId27 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            RunFonts runFonts106 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize82 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript82 = new FontSizeComplexScript() { Val = "26" };
            Languages languages45 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties29.Append(runFonts106);
            paragraphMarkRunProperties29.Append(fontSize82);
            paragraphMarkRunProperties29.Append(fontSizeComplexScript82);
            paragraphMarkRunProperties29.Append(languages45);

            paragraphProperties29.Append(paragraphStyleId27);
            paragraphProperties29.Append(paragraphMarkRunProperties29);

            Run run110 = new Run();

            RunProperties runProperties80 = new RunProperties();
            RunFonts runFonts107 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize83 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript83 = new FontSizeComplexScript() { Val = "26" };
            Languages languages46 = new Languages() { Val = "en-US" };

            runProperties80.Append(runFonts107);
            runProperties80.Append(fontSize83);
            runProperties80.Append(fontSizeComplexScript83);
            runProperties80.Append(languages46);
            Text text54 = new Text();
            text54.Text = "Maths";

            run110.Append(runProperties80);
            run110.Append(text54);

            paragraph31.Append(paragraphProperties29);
            paragraph31.Append(run110);

            tableCell26.Append(tableCellProperties26);
            tableCell26.Append(paragraph31);

            TableCell tableCell27 = new TableCell();

            TableCellProperties tableCellProperties27 = new TableCellProperties();
            TableCellWidth tableCellWidth27 = new TableCellWidth() { Width = "1607", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge6 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders27 = new TableCellBorders();
            LeftBorder leftBorder28 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder28 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            RightBorder rightBorder7 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder28 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder7 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders27.Append(leftBorder28);
            tableCellBorders27.Append(bottomBorder28);
            tableCellBorders27.Append(rightBorder7);
            tableCellBorders27.Append(insideHorizontalBorder28);
            tableCellBorders27.Append(insideVerticalBorder7);
            Shading shading27 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin27 = new TableCellMargin();
            LeftMargin leftMargin27 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin27.Append(leftMargin27);
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties27.Append(tableCellWidth27);
            tableCellProperties27.Append(verticalMerge6);
            tableCellProperties27.Append(tableCellBorders27);
            tableCellProperties27.Append(shading27);
            tableCellProperties27.Append(tableCellMargin27);
            tableCellProperties27.Append(tableCellVerticalAlignment2);

            Paragraph paragraph32 = new Paragraph();

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId28 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            RunFonts runFonts108 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize84 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript84 = new FontSizeComplexScript() { Val = "26" };
            Languages languages47 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties30.Append(runFonts108);
            paragraphMarkRunProperties30.Append(fontSize84);
            paragraphMarkRunProperties30.Append(fontSizeComplexScript84);
            paragraphMarkRunProperties30.Append(languages47);

            paragraphProperties30.Append(paragraphStyleId28);
            paragraphProperties30.Append(paragraphMarkRunProperties30);

            Run run111 = new Run();

            RunProperties runProperties81 = new RunProperties();
            RunFonts runFonts109 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize85 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript85 = new FontSizeComplexScript() { Val = "26" };
            Languages languages48 = new Languages() { Val = "en-US" };

            runProperties81.Append(runFonts109);
            runProperties81.Append(fontSize85);
            runProperties81.Append(fontSizeComplexScript85);
            runProperties81.Append(languages48);
            Text text55 = new Text();
            text55.Text = "Project";

            run111.Append(runProperties81);
            run111.Append(text55);

            paragraph32.Append(paragraphProperties30);
            paragraph32.Append(run111);

            tableCell27.Append(tableCellProperties27);
            tableCell27.Append(paragraph32);

            tableRow6.Append(tableRowProperties6);
            tableRow6.Append(tableCell22);
            tableRow6.Append(tableCell23);
            tableRow6.Append(tableCell24);
            tableRow6.Append(tableCell25);
            tableRow6.Append(tableCell26);
            tableRow6.Append(tableCell27);

            TableRow tableRow7 = new TableRow();
            TableRowProperties tableRowProperties7 = new TableRowProperties();

            TableCell tableCell28 = new TableCell();

            TableCellProperties tableCellProperties28 = new TableCellProperties();
            TableCellWidth tableCellWidth28 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge7 = new VerticalMerge() { Val = MergedCellValues.Continue };

            TableCellBorders tableCellBorders28 = new TableCellBorders();
            LeftBorder leftBorder29 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder29 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder29 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders28.Append(leftBorder29);
            tableCellBorders28.Append(bottomBorder29);
            tableCellBorders28.Append(insideHorizontalBorder29);
            Shading shading28 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin28 = new TableCellMargin();
            LeftMargin leftMargin28 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin28.Append(leftMargin28);

            tableCellProperties28.Append(tableCellWidth28);
            tableCellProperties28.Append(verticalMerge7);
            tableCellProperties28.Append(tableCellBorders28);
            tableCellProperties28.Append(shading28);
            tableCellProperties28.Append(tableCellMargin28);

            Paragraph paragraph33 = new Paragraph();

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId29 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            RunFonts runFonts110 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize86 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript86 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties31.Append(runFonts110);
            paragraphMarkRunProperties31.Append(fontSize86);
            paragraphMarkRunProperties31.Append(fontSizeComplexScript86);

            paragraphProperties31.Append(paragraphStyleId29);
            paragraphProperties31.Append(paragraphMarkRunProperties31);

            Run run112 = new Run();

            RunProperties runProperties82 = new RunProperties();
            RunFonts runFonts111 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize87 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript87 = new FontSizeComplexScript() { Val = "26" };

            runProperties82.Append(runFonts111);
            runProperties82.Append(fontSize87);
            runProperties82.Append(fontSizeComplexScript87);

            run112.Append(runProperties82);

            paragraph33.Append(paragraphProperties31);
            paragraph33.Append(run112);

            tableCell28.Append(tableCellProperties28);
            tableCell28.Append(paragraph33);

            TableCell tableCell29 = new TableCell();

            TableCellProperties tableCellProperties29 = new TableCellProperties();
            TableCellWidth tableCellWidth29 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders29 = new TableCellBorders();
            LeftBorder leftBorder30 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder30 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder30 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders29.Append(leftBorder30);
            tableCellBorders29.Append(bottomBorder30);
            tableCellBorders29.Append(insideHorizontalBorder30);
            Shading shading29 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin29 = new TableCellMargin();
            LeftMargin leftMargin29 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin29.Append(leftMargin29);

            tableCellProperties29.Append(tableCellWidth29);
            tableCellProperties29.Append(tableCellBorders29);
            tableCellProperties29.Append(shading29);
            tableCellProperties29.Append(tableCellMargin29);

            Paragraph paragraph34 = new Paragraph();

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId30 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            RunFonts runFonts112 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize88 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript88 = new FontSizeComplexScript() { Val = "26" };
            Languages languages49 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties32.Append(runFonts112);
            paragraphMarkRunProperties32.Append(fontSize88);
            paragraphMarkRunProperties32.Append(fontSizeComplexScript88);
            paragraphMarkRunProperties32.Append(languages49);

            paragraphProperties32.Append(paragraphStyleId30);
            paragraphProperties32.Append(paragraphMarkRunProperties32);

            Run run113 = new Run();

            RunProperties runProperties83 = new RunProperties();
            RunFonts runFonts113 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize89 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript89 = new FontSizeComplexScript() { Val = "26" };
            Languages languages50 = new Languages() { Val = "en-US" };

            runProperties83.Append(runFonts113);
            runProperties83.Append(fontSize89);
            runProperties83.Append(fontSizeComplexScript89);
            runProperties83.Append(languages50);
            Text text56 = new Text();
            text56.Text = "Social";

            run113.Append(runProperties83);
            run113.Append(text56);

            paragraph34.Append(paragraphProperties32);
            paragraph34.Append(run113);

            tableCell29.Append(tableCellProperties29);
            tableCell29.Append(paragraph34);

            TableCell tableCell30 = new TableCell();

            TableCellProperties tableCellProperties30 = new TableCellProperties();
            TableCellWidth tableCellWidth30 = new TableCellWidth() { Width = "1607", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders30 = new TableCellBorders();
            LeftBorder leftBorder31 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder31 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder31 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders30.Append(leftBorder31);
            tableCellBorders30.Append(bottomBorder31);
            tableCellBorders30.Append(insideHorizontalBorder31);
            Shading shading30 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin30 = new TableCellMargin();
            LeftMargin leftMargin30 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin30.Append(leftMargin30);

            tableCellProperties30.Append(tableCellWidth30);
            tableCellProperties30.Append(tableCellBorders30);
            tableCellProperties30.Append(shading30);
            tableCellProperties30.Append(tableCellMargin30);

            Paragraph paragraph35 = new Paragraph();

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId31 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            RunFonts runFonts114 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize90 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript90 = new FontSizeComplexScript() { Val = "26" };
            Languages languages51 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties33.Append(runFonts114);
            paragraphMarkRunProperties33.Append(fontSize90);
            paragraphMarkRunProperties33.Append(fontSizeComplexScript90);
            paragraphMarkRunProperties33.Append(languages51);

            paragraphProperties33.Append(paragraphStyleId31);
            paragraphProperties33.Append(paragraphMarkRunProperties33);

            Run run114 = new Run();

            RunProperties runProperties84 = new RunProperties();
            RunFonts runFonts115 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize91 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript91 = new FontSizeComplexScript() { Val = "26" };
            Languages languages52 = new Languages() { Val = "en-US" };

            runProperties84.Append(runFonts115);
            runProperties84.Append(fontSize91);
            runProperties84.Append(fontSizeComplexScript91);
            runProperties84.Append(languages52);
            Text text57 = new Text();
            text57.Text = "History";

            run114.Append(runProperties84);
            run114.Append(text57);

            paragraph35.Append(paragraphProperties33);
            paragraph35.Append(run114);

            tableCell30.Append(tableCellProperties30);
            tableCell30.Append(paragraph35);

            TableCell tableCell31 = new TableCell();

            TableCellProperties tableCellProperties31 = new TableCellProperties();
            TableCellWidth tableCellWidth31 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders31 = new TableCellBorders();
            LeftBorder leftBorder32 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder32 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder32 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders31.Append(leftBorder32);
            tableCellBorders31.Append(bottomBorder32);
            tableCellBorders31.Append(insideHorizontalBorder32);
            Shading shading31 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin31 = new TableCellMargin();
            LeftMargin leftMargin31 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin31.Append(leftMargin31);

            tableCellProperties31.Append(tableCellWidth31);
            tableCellProperties31.Append(tableCellBorders31);
            tableCellProperties31.Append(shading31);
            tableCellProperties31.Append(tableCellMargin31);

            Paragraph paragraph36 = new Paragraph();

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId32 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            RunFonts runFonts116 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize92 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript92 = new FontSizeComplexScript() { Val = "26" };
            Languages languages53 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties34.Append(runFonts116);
            paragraphMarkRunProperties34.Append(fontSize92);
            paragraphMarkRunProperties34.Append(fontSizeComplexScript92);
            paragraphMarkRunProperties34.Append(languages53);

            paragraphProperties34.Append(paragraphStyleId32);
            paragraphProperties34.Append(paragraphMarkRunProperties34);

            Run run115 = new Run();

            RunProperties runProperties85 = new RunProperties();
            RunFonts runFonts117 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize93 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript93 = new FontSizeComplexScript() { Val = "26" };
            Languages languages54 = new Languages() { Val = "en-US" };

            runProperties85.Append(runFonts117);
            runProperties85.Append(fontSize93);
            runProperties85.Append(fontSizeComplexScript93);
            runProperties85.Append(languages54);
            Text text58 = new Text();
            text58.Text = "English";

            run115.Append(runProperties85);
            run115.Append(text58);

            paragraph36.Append(paragraphProperties34);
            paragraph36.Append(run115);

            tableCell31.Append(tableCellProperties31);
            tableCell31.Append(paragraph36);

            TableCell tableCell32 = new TableCell();

            TableCellProperties tableCellProperties32 = new TableCellProperties();
            TableCellWidth tableCellWidth32 = new TableCellWidth() { Width = "1606", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders32 = new TableCellBorders();
            LeftBorder leftBorder33 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder33 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder33 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders32.Append(leftBorder33);
            tableCellBorders32.Append(bottomBorder33);
            tableCellBorders32.Append(insideHorizontalBorder33);
            Shading shading32 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin32 = new TableCellMargin();
            LeftMargin leftMargin32 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin32.Append(leftMargin32);

            tableCellProperties32.Append(tableCellWidth32);
            tableCellProperties32.Append(tableCellBorders32);
            tableCellProperties32.Append(shading32);
            tableCellProperties32.Append(tableCellMargin32);

            Paragraph paragraph37 = new Paragraph();

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId33 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            RunFonts runFonts118 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize94 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript94 = new FontSizeComplexScript() { Val = "26" };
            Languages languages55 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties35.Append(runFonts118);
            paragraphMarkRunProperties35.Append(fontSize94);
            paragraphMarkRunProperties35.Append(fontSizeComplexScript94);
            paragraphMarkRunProperties35.Append(languages55);

            paragraphProperties35.Append(paragraphStyleId33);
            paragraphProperties35.Append(paragraphMarkRunProperties35);

            Run run116 = new Run();

            RunProperties runProperties86 = new RunProperties();
            RunFonts runFonts119 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize95 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript95 = new FontSizeComplexScript() { Val = "26" };
            Languages languages56 = new Languages() { Val = "en-US" };

            runProperties86.Append(runFonts119);
            runProperties86.Append(fontSize95);
            runProperties86.Append(fontSizeComplexScript95);
            runProperties86.Append(languages56);
            Text text59 = new Text();
            text59.Text = "Social";

            run116.Append(runProperties86);
            run116.Append(text59);

            paragraph37.Append(paragraphProperties35);
            paragraph37.Append(run116);

            tableCell32.Append(tableCellProperties32);
            tableCell32.Append(paragraph37);

            TableCell tableCell33 = new TableCell();

            TableCellProperties tableCellProperties33 = new TableCellProperties();
            TableCellWidth tableCellWidth33 = new TableCellWidth() { Width = "1607", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge8 = new VerticalMerge() { Val = MergedCellValues.Continue };

            TableCellBorders tableCellBorders33 = new TableCellBorders();
            LeftBorder leftBorder34 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder34 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            RightBorder rightBorder8 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder34 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder8 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders33.Append(leftBorder34);
            tableCellBorders33.Append(bottomBorder34);
            tableCellBorders33.Append(rightBorder8);
            tableCellBorders33.Append(insideHorizontalBorder34);
            tableCellBorders33.Append(insideVerticalBorder8);
            Shading shading33 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

            TableCellMargin tableCellMargin33 = new TableCellMargin();
            LeftMargin leftMargin33 = new LeftMargin() { Width = "54", Type = TableWidthUnitValues.Dxa };

            tableCellMargin33.Append(leftMargin33);

            tableCellProperties33.Append(tableCellWidth33);
            tableCellProperties33.Append(verticalMerge8);
            tableCellProperties33.Append(tableCellBorders33);
            tableCellProperties33.Append(shading33);
            tableCellProperties33.Append(tableCellMargin33);

            Paragraph paragraph38 = new Paragraph();

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId34 = new ParagraphStyleId() { Val = "TableContents" };

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            RunFonts runFonts120 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize96 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript96 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties36.Append(runFonts120);
            paragraphMarkRunProperties36.Append(fontSize96);
            paragraphMarkRunProperties36.Append(fontSizeComplexScript96);

            paragraphProperties36.Append(paragraphStyleId34);
            paragraphProperties36.Append(paragraphMarkRunProperties36);

            Run run117 = new Run();

            RunProperties runProperties87 = new RunProperties();
            RunFonts runFonts121 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            FontSize fontSize97 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript97 = new FontSizeComplexScript() { Val = "26" };

            runProperties87.Append(runFonts121);
            runProperties87.Append(fontSize97);
            runProperties87.Append(fontSizeComplexScript97);

            run117.Append(runProperties87);

            paragraph38.Append(paragraphProperties36);
            paragraph38.Append(run117);

            tableCell33.Append(tableCellProperties33);
            tableCell33.Append(paragraph38);

            tableRow7.Append(tableRowProperties7);
            tableRow7.Append(tableCell28);
            tableRow7.Append(tableCell29);
            tableRow7.Append(tableCell30);
            tableRow7.Append(tableCell31);
            tableRow7.Append(tableCell32);
            tableRow7.Append(tableCell33);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);
            table1.Append(tableRow5);
            table1.Append(tableRow6);
            table1.Append(tableRow7);
            //body.InsertBefore(table1, sectionProperties1);

            Paragraph paragraph39 = new Paragraph();

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId35 = new ParagraphStyleId() { Val = "Normal" };
            WidowControl widowControl2 = new WidowControl();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation2 = new Indentation() { Left = "0", Right = "0", Hanging = "0" };
            Justification justification10 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            RunFonts runFonts122 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
            FontSize fontSize98 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript98 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties37.Append(runFonts122);
            paragraphMarkRunProperties37.Append(fontSize98);
            paragraphMarkRunProperties37.Append(fontSizeComplexScript98);

            paragraphProperties37.Append(paragraphStyleId35);
            paragraphProperties37.Append(widowControl2);
            paragraphProperties37.Append(spacingBetweenLines2);
            paragraphProperties37.Append(indentation2);
            paragraphProperties37.Append(justification10);
            paragraphProperties37.Append(paragraphMarkRunProperties37);

            Run run118 = new Run();
            RunProperties runProperties88 = new RunProperties();

            run118.Append(runProperties88);

            paragraph39.Append(paragraphProperties37);
            paragraph39.Append(run118);
            //body.InsertBefore(paragraph39, sectionProperties1);

            //    DocGrid docGrid1 = sectionProperties1.GetFirstChild<DocGrid>();
            //docGrid1.CharacterSpace = new Int32Value() { InnerText = "4294961151" };
        }

        //private void ChangeStyleDefinitionsPart1(StyleDefinitionsPart styleDefinitionsPart1)
        //{
        //    Styles styles1 = styleDefinitionsPart1.Styles;

        //    DocDefaults docDefaults1 = styles1.GetFirstChild<DocDefaults>();
        //    Style style1 = styles1.GetFirstChild<Style>();
        //    Style style2 = styles1.Elements<Style>().ElementAt(1);
        //    Style style3 = styles1.Elements<Style>().ElementAt(2);
        //    Style style4 = styles1.Elements<Style>().ElementAt(3);
        //    Style style5 = styles1.Elements<Style>().ElementAt(9);
        //    Style style6 = styles1.Elements<Style>().ElementAt(13);
        //    Style style7 = styles1.Elements<Style>().ElementAt(14);

        //    RunPropertiesDefault runPropertiesDefault1 = docDefaults1.GetFirstChild<RunPropertiesDefault>();

        //    RunPropertiesBaseStyle runPropertiesBaseStyle1 = runPropertiesDefault1.GetFirstChild<RunPropertiesBaseStyle>();

        //    FontSizeComplexScript fontSizeComplexScript1 = runPropertiesBaseStyle1.GetFirstChild<FontSizeComplexScript>();

        //    FontSize fontSize1 = new FontSize() { Val = "20" };
        //    runPropertiesBaseStyle1.InsertBefore(fontSize1, fontSizeComplexScript1);

        //    StyleParagraphProperties styleParagraphProperties1 = style1.GetFirstChild<StyleParagraphProperties>();
        //    StyleRunProperties styleRunProperties1 = style1.GetFirstChild<StyleRunProperties>();

        //    Justification justification1 = new Justification() { Val = JustificationValues.Left };
        //    styleParagraphProperties1.Append(justification1);

        //    Color color1 = styleRunProperties1.GetFirstChild<Color>();
        //    color1.Val = "00000A";

        //    NextParagraphStyle nextParagraphStyle1 = style2.GetFirstChild<NextParagraphStyle>();
        //    StyleParagraphProperties styleParagraphProperties2 = style2.GetFirstChild<StyleParagraphProperties>();

        //    nextParagraphStyle1.Remove();

        //    NumberingProperties numberingProperties1 = styleParagraphProperties2.GetFirstChild<NumberingProperties>();
        //    OutlineLevel outlineLevel1 = styleParagraphProperties2.Elements<OutlineLevel>().ElementAt(1);

        //    numberingProperties1.Remove();
        //    outlineLevel1.Remove();

        //    NextParagraphStyle nextParagraphStyle2 = style3.GetFirstChild<NextParagraphStyle>();
        //    StyleParagraphProperties styleParagraphProperties3 = style3.GetFirstChild<StyleParagraphProperties>();

        //    nextParagraphStyle2.Remove();

        //    NumberingProperties numberingProperties2 = styleParagraphProperties3.GetFirstChild<NumberingProperties>();
        //    OutlineLevel outlineLevel2 = styleParagraphProperties3.Elements<OutlineLevel>().ElementAt(1);

        //    numberingProperties2.Remove();
        //    outlineLevel2.Remove();

        //    NextParagraphStyle nextParagraphStyle3 = style4.GetFirstChild<NextParagraphStyle>();
        //    StyleParagraphProperties styleParagraphProperties4 = style4.GetFirstChild<StyleParagraphProperties>();

        //    nextParagraphStyle3.Remove();

        //    NumberingProperties numberingProperties3 = styleParagraphProperties4.GetFirstChild<NumberingProperties>();
        //    OutlineLevel outlineLevel3 = styleParagraphProperties4.Elements<OutlineLevel>().ElementAt(1);

        //    numberingProperties3.Remove();
        //    outlineLevel3.Remove();

        //    NextParagraphStyle nextParagraphStyle4 = style5.GetFirstChild<NextParagraphStyle>();

        //    nextParagraphStyle4.Remove();

        //    NextParagraphStyle nextParagraphStyle5 = style6.GetFirstChild<NextParagraphStyle>();

        //    nextParagraphStyle5.Remove();

        //    NextParagraphStyle nextParagraphStyle6 = style7.GetFirstChild<NextParagraphStyle>();

        //    nextParagraphStyle6.Remove();

        //    Style style8 = new Style() { Type = StyleValues.Paragraph, StyleId = "TableContents" };
        //    StyleName styleName1 = new StyleName() { Val = "Table Contents" };
        //    BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
        //    PrimaryStyle primaryStyle1 = new PrimaryStyle();
        //    StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
        //    StyleRunProperties styleRunProperties2 = new StyleRunProperties();

        //    style8.Append(styleName1);
        //    style8.Append(basedOn1);
        //    style8.Append(primaryStyle1);
        //    style8.Append(styleParagraphProperties5);
        //    style8.Append(styleRunProperties2);
        //    styles1.Append(style8);

        //    Style style9 = new Style() { Type = StyleValues.Paragraph, StyleId = "TableHeading" };
        //    StyleName styleName2 = new StyleName() { Val = "Table Heading" };
        //    BasedOn basedOn2 = new BasedOn() { Val = "TableContents" };
        //    PrimaryStyle primaryStyle2 = new PrimaryStyle();
        //    StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();
        //    StyleRunProperties styleRunProperties3 = new StyleRunProperties();

        //    style9.Append(styleName2);
        //    style9.Append(basedOn2);
        //    style9.Append(primaryStyle2);
        //    style9.Append(styleParagraphProperties6);
        //    style9.Append(styleRunProperties3);
        //    styles1.Append(style9);
        //}
    }
}
