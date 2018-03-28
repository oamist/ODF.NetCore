using ODFLib.Document.Content.Draw;
using ODFLib.Document.Content.Tables;
using ODFLib.Document.Content.Text;
using ODFLib.Document.Styles;
using ODFLib.Document.TextDocuments;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestConsoleApp
{
    public class Example02
    {
        /// <summary>
        ///  3.Add Image Maps to your document.
        /// </summary>
        public static void Test03()
        {
            string imagePath = @".\Pictures\img01.jpg";
            TextDocument document = new TextDocument();
            document.New();
            //Create standard paragraph
            Paragraph paragraphOuter = ParagraphBuilder.CreateStandardTextParagraph(document);
            //Create the frame with graphic
            Frame frame = new Frame(document, "frame1", "graphic1", imagePath);
            //Create a Draw Area Rectangle
            DrawAreaRectangle drawAreaRec = new DrawAreaRectangle(
                document, "0cm", "0cm", "1.5cm", "2.5cm", null);
            drawAreaRec.Href = "http://www.ncwu.edu.cn/";
            //Create a Draw Area Circle
            DrawAreaCircle drawAreaCircle = new DrawAreaCircle(
                document, "4cm", "4cm", "1.5cm", null);
            drawAreaCircle.Href = "http://www.ncwu.edu.cn/";
            DrawArea[] drawArea = new DrawArea[2] { drawAreaRec, drawAreaCircle };
            //Create a Image Map
            ImageMap imageMap = new ImageMap(document, drawArea);
            //Add Image Map to the frame
            frame.Content.Add(imageMap);
            //Add frame to paragraph
            paragraphOuter.Content.Add(frame);
            //Add paragraph to document
            document.Content.Add(paragraphOuter);
            //Save the document
            document.SaveTo("simpleImageMap.odt");
        }

        /// <summary>
        ///   4.Add a graphic as Illustration to you document.
        /// </summary>
        public static void Test04()
        {
            TextDocument textdocument = new TextDocument();
            textdocument.New();
            Paragraph pOuter = ParagraphBuilder.CreateStandardTextParagraph(textdocument);
            DrawTextBox drawTextBox = new DrawTextBox(textdocument);
            Frame frameTextBox = new Frame(textdocument, "fr_txt_box");
            frameTextBox.DrawName = "fr_txt_box";
            frameTextBox.ZIndex = "0";
            Paragraph p = ParagraphBuilder.CreateStandardTextParagraph(textdocument);
            p.StyleName = "Illustration";
            Frame frame = new Frame(textdocument, "frame1", "graphic1", @".\Pictures\img02.jpg");
            frame.ZIndex = "1";
            p.Content.Add(frame);
            p.TextContent.Add(new SimpleText(textdocument, "Illustration"));
            drawTextBox.Content.Add(p);
            frameTextBox.SvgWidth = frame.SvgWidth;
            drawTextBox.MinWidth = frame.SvgWidth;
            drawTextBox.MinHeight = frame.SvgHeight;
            frameTextBox.Content.Add(drawTextBox);
            pOuter.Content.Add(frameTextBox);
            textdocument.Content.Add(pOuter);
            textdocument.SaveTo("drawTextbox.odt");
        }

        /// <summary>
        ///  5.Add graphics to your document.
        /// </summary>
        public static void Test05()
        {
            TextDocument textdocument = new TextDocument();
            textdocument.New();
            Paragraph p = ParagraphBuilder.CreateStandardTextParagraph(textdocument);
            Frame frame = new Frame(textdocument, "frame1", "graphic1", @".\Pictures\img03.jpg");
            p.Content.Add(frame);
            textdocument.Content.Add(p);
            textdocument.SaveTo("grapic.odt");
        }

        /// <summary>
        ///   6.Create a table and use cell merging.
        /// </summary>
        public static void Test06()
        {
            //Create a new text document
            TextDocument document = new TextDocument();
            document.New();
            //Create a table for a text document using the TableBuilder
            Table table = TableBuilder.CreateTextDocumentTable(
                document,
                "table1",
                "table1",
                3,
                3,
                16.99,
                false,
                false);

            //Fill the cells
            foreach (Row row in table.RowCollection)
            {
                foreach (Cell cell in row.CellCollection)
                {
                    //Create a standard paragraph
                    Paragraph paragraph = ParagraphBuilder.CreateStandardTextParagraph(document);
                    //Add some simple text
                    paragraph.TextContent.Add(new SimpleText(document, "Cell text"));
                    cell.Content.Add(paragraph);
                }
            }
            //Merge some cells. Notice this is only available in text documents!
            table.RowCollection[1].MergeCells(document, 1, 2, true);
            //Add table to the document
            document.Content.Add(table);
            //Save the document
            document.SaveTo("simpleTableWithMergedCells.odt");
        }
        /// <summary>
        ///  7.Create a simple text table by using the TableBuilder class.
        /// </summary>
        public static void Test07()
        {
            //Create a new text document
            TextDocument document = new TextDocument();
            document.New();
            //Create a table for a text document using the TableBuilder
            Table table = TableBuilder.CreateTextDocumentTable(
                document,
                "table1",
                "table1",
                3,
                3,
                16.99,
                false,
                false);
            //Create a standard paragraph
            Paragraph paragraph = ParagraphBuilder.CreateStandardTextParagraph(document);
            //Add some simple text
            paragraph.TextContent.Add(new SimpleText(document, "Some cell text"));
            //Insert paragraph into the first cell
            table.RowCollection[0].CellCollection[0].Content.Add(paragraph);
            //Add table to the document
            document.Content.Add(table);
            //Save the document
            document.SaveTo("simpleTable.odt");
        }
        /// <summary>
        ///  8.Don't do things twice. Create deep clones of any IContent object.
        /// </summary>
        public static void Test08()
        {
            TextDocument document = new TextDocument();
            document.New();
            Paragraph paragraph = ParagraphBuilder.CreateStandardTextParagraph(document);
            paragraph.TextContent.Add(new SimpleText(document, "Some text"));
            Paragraph paragraphClone = (Paragraph)paragraph.Clone();
            ParagraphStyle paragraphStyle = new ParagraphStyle(document, "P1");
            paragraphStyle.TextProperties.Bold = "bold";
            //Add paragraph style to the document, 
            //only automaticaly created styles will be added also automaticaly
            document.Styles.Add(paragraphStyle);
            paragraphClone.ParagraphStyle = paragraphStyle;
            //Clone the clone
            Paragraph paragraphClone2 = (Paragraph)paragraphClone.Clone();
            document.Content.Add(paragraph);
            document.Content.Add(paragraphClone);
            document.Content.Add(paragraphClone2);
            document.SaveTo("clonedParagraphs.odt");
        }
        /// <summary>
        ///  9.Accessing and manipulating common styles (style templates) 
        /// </summary>
        public static void Test09()
        {
            TextDocument document = new TextDocument();
            document.New();
            //Find a Header template
            IStyle style = document.CommonStyles.GetStyleByName("Heading_20_1");
            //Assert.IsNotNull(style, "Style with name Heading_20_1 must exist");
            //Assert.IsTrue(style is ParagraphStyle, "style must be a ParagraphStyle");
            ((ParagraphStyle)style).TextProperties.FontName = FontFamilies.BroadwayBT;
            //Create a header that use the standard style Heading_20_1
            Header header = new Header(document, Headings.Heading_20_1);
            //Add some text
            header.TextContent.Add(new SimpleText(document, "I am the header text and my style template was modified :)"));
            //Add header to the document
            document.Content.Add(header);
            //save the document
            document.SaveTo("modifiedCommonStyle.odt");

        }

        /// <summary>
        ///  10.Create a Paragraph Collection from a long string by using the ParagraphBuilder.
        /// </summary>
        public static void Test10()
        {
            //some text e.g read from a TextBox
            string someText = "Max Mustermann\nMustermann Str. 300\n22222 Hamburg\n\n\n\n"
                                    + "Heinz Willi\nDorfstr. 1\n22225 Hamburg\n\n\n\n"
                                    + "Offer for 200 Intel Pentium 4 CPU's\n\n\n\n"
                                    + "Dear Mr. Willi,\n\n\n\n"
                                    + "thank you for your request. \tWe can "
                                    + "offer you the 200 Intel Pentium IV 3 Ghz CPU's for a price of "
                                    + "79,80 € per unit."
                                    + "This special offer is valid to 31.10.2005. If you accept, we "
                                    + "can deliver within 24 hours.\n\n\n\n"
                                    + "Best regards \nMax Mustermann";

            //Create new TextDocument
            TextDocument document = new TextDocument();
            document.New();
            //Use the ParagraphBuilder to split the string into ParagraphCollection
            ParagraphCollection pCollection = ParagraphBuilder.CreateParagraphCollection(
                                                 document,
                                                 someText,
                                                 true,
                                                 ParagraphBuilder.ParagraphSeperator);
            //Add the paragraph collection
            foreach (Paragraph paragraph in pCollection)
                document.Content.Add(paragraph);
            //save
            document.SaveTo("Letter.odt");
        }
        /// <summary>
        ///   11.Creating Footnotes and Endnotes.
        /// </summary>
        public static void Test11()
        {
            //Create new TextDocument
            TextDocument document = new TextDocument();
            document.New();
            //Create a new Paragph
            Paragraph para = ParagraphBuilder.CreateStandardTextParagraph(document);
            //Create some simple Text
            para.TextContent.Add(new SimpleText(document, "Some simple text. And I have a footnode"));
            //Create a Footnote
            para.TextContent.Add(new Footnote(document, "Footer Text", "1", FootnoteType.footnode));
            //Add the paragraph
            document.Content.Add(para);
            //Save
            document.SaveTo("footnote.odt");
        }

        /// <summary>
        /// 12.Create any type of hyperlinks by using the XLink class 
        /// </summary>
        public static void Test12()
        {
            // Create new TextDocument
            TextDocument document = new TextDocument();
            document.New();
            //Create a new Paragraph
            Paragraph para = new Paragraph(document, "P1");
            //Create some simple text
            SimpleText stext = new SimpleText(document, "Some simple text ");
            //Create a XLink
            XLink xlink = new XLink(document, "http://www.ncwu.edu.cn/", "ncwu.edu.cn");
            //Add the textcontent
            para.TextContent.Add(stext);
            para.TextContent.Add(xlink);
            //Add paragraph to the document content
            document.Content.Add(para);
            //Save
            document.SaveTo("XLink.odt");
        }

        /// <summary>
        ///  13.Using the Header object and fill the heading text using the TextBuilder class.
        /// </summary>
        public static void Test13()
        {
            string headingText = "Some    Heading with\n styles\t,line breaks, tab stops and extra whitspaces";
            //Create a new text document
            TextDocument document = new TextDocument();
            document.New();
            //Create a new Heading
            Header header = new Header(document, Headings.Heading);
            //Create a TextCollection from headingText using the TextBuilder
            //You can conert every string incl. control character \n, \t, .. into
            //a ItextCollection using the TextBuilder
            ITextCollection textCol = TextBuilder.BuildTextCollection(document, headingText);
            //Add text collection
            header.TextContent = textCol;
            //Add header
            document.Content.Add(header);
            document.SaveTo("HeadingWithControlCharacter.odt");
        }
        /// <summary>
        ///  14.Create a List with List Item objects.
        /// </summary>
        public static void Test14()
        {
            //Create a new text document
            TextDocument document = new TextDocument();
            document.New();
            //Create a numbered list
            List li = new List(document, "L1", ListStyles.Bullet, "L1P1");
            //Create a new list item
            ListItem lit = new ListItem(li);
            //Create a paragraph	
            Paragraph paragraph = ParagraphBuilder.CreateStandardTextParagraph(document);
            //Add some text
            paragraph.TextContent.Add(new SimpleText(document, "First item"));
            //Add paragraph to the list item
            lit.Content.Add(paragraph);
            //Add the list item
            li.Content.Add(lit);
            //Add the list
            document.Content.Add(li);
            //Save document
            document.SaveTo("list.odt");
        }
        /// <summary>
        ///   15.Create a standard Paragraph with formated text.
        /// </summary>
        public static void Test15()
        {
            //Create a new text document
            TextDocument document = new TextDocument();
            document.New();
            //Create a standard paragraph using the ParagraphBuilder
            Paragraph paragraph = ParagraphBuilder.CreateStandardTextParagraph(document);
            //Add some formated text
            FormatedText formText = new FormatedText(document, "T1", "Some formated text!");
            formText.TextStyle.TextProperties.Bold = "bold";
            paragraph.TextContent.Add(formText);
            //Add the paragraph to the document
            document.Content.Add(paragraph);
            //Save 
            document.SaveTo("formated.odt");
        }
        /// <summary>
        ///   16.Create a simple Paragraph using the ParagraphBuilder.
        /// </summary>
        public static void Test16()
        {
            //Create a new text document
            TextDocument document = new TextDocument();
            document.New();
            //Create a standard paragraph using the ParagraphBuilder
            Paragraph paragraph = ParagraphBuilder.CreateStandardTextParagraph(document);
            //Add some simple text
            paragraph.TextContent.Add(new SimpleText(document, "Some simple text!"));
            //Add the paragraph to the document
            document.Content.Add(paragraph);
            //Save 
            document.SaveTo("simple.odt");
        }
       
    }
}
