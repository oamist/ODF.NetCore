using ODFLib.Document.Content.Tables;
using ODFLib.Document.Content.Text;
using ODFLib.Document.SpreadsheetDocuments;
using ODFLib.Document.Styles;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestConsoleApp
{
    public class Example01
    {
        /// <summary>
        ///1. Create a Spreadsheet document and add a table with formated cells. 
        /// </summary>
        public static void Test01()
        {
            //Create new spreadsheet document
            SpreadsheetDocument spreadsheetDocument = new SpreadsheetDocument();
            spreadsheetDocument.New();
            //Create a new table
            Table table = new Table(spreadsheetDocument, "First", "tablefirst");
            //Create a new cell, without any extra styles 
            Cell cell = table.CreateCell();
            //Add a paragraph to this cell
            Paragraph paragraph = ParagraphBuilder.CreateSpreadsheetParagraph(spreadsheetDocument);
            //Create some Formated text
            FormatedText fText = new FormatedText(spreadsheetDocument, "T1", "Some Text");
            //fText.TextStyle.TextProperties.Bold = "bold";
            fText.TextStyle.TextProperties.Underline = LineStyles.dotted;
            //Add formated text
            paragraph.TextContent.Add(fText);
            //Add paragraph to the cell
            cell.Content.Add(paragraph);
            //Insert the cell at row index 1 and column index 2
            //All need rows, columns and cells below the given
            //indexes will be build automatically.
            table.InsertCellAt(1, 2, cell);
            //Insert table into the spreadsheet document
            spreadsheetDocument.TableCollection.Add(table);
            spreadsheetDocument.SaveTo("formated.ods");
        }

        //2.Create a spreadsheet document and add a table 
        public static void Test02()
        {
            //Create new spreadsheet document
            SpreadsheetDocument spreadsheetDocument = new SpreadsheetDocument();
            spreadsheetDocument.New();
            //Create a new table
            Table table = new Table(spreadsheetDocument, "First", "tablefirst");
            //Create a new cell, without any extra styles 
            Cell cell = table.CreateCell("cell001");
            cell.OfficeValueType = "string";
            //Set full border
            cell.CellStyle.CellProperties.Border = Border.NormalSolid;
            //Add a paragraph to this cell
            Paragraph paragraph = ParagraphBuilder.CreateSpreadsheetParagraph(spreadsheetDocument);
            //Add some text content
            paragraph.TextContent.Add(new SimpleText(spreadsheetDocument, "Some text"));
            //Add paragraph to the cell
            cell.Content.Add(paragraph);
            //Insert the cell at row index 1 and column index 2
            //All need rows, columns and cells below the given
            //indexes will be build automatically.
            table.InsertCellAt(1, 2, cell);
            //Insert table into the spreadsheet document
            spreadsheetDocument.TableCollection.Add(table);
            spreadsheetDocument.SaveTo("simple.ods");
        }

    }
}
