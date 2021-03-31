using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GenerateWordDoc
{
    class Program
    {
        private class DocumentBuilder
        {
            private readonly WordprocessingDocument _document;
            private readonly MainDocumentPart _mainPart;
            public readonly Body Body;

            public DocumentBuilder(string filePath)
            {
                _document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
                _mainPart = _document.AddMainDocumentPart();
                _mainPart.Document = new Document(new Body());
                Body = _mainPart.Document.Body;
            }

            public void AddParagraph(string text)
            {
                var run = new Run(new Text(text));
                var runProperties = new RunProperties();
                runProperties.AppendChild(new Bold());
                run.RunProperties = runProperties;

                var para = new Paragraph(run);
                Body.AppendChild(para);
            }

            public void AddTable(List<List<string>> data)
            {
                var table = new Table();
                data.ForEach(dataRow =>
                {
                    var tableRow = new TableRow();
                    dataRow.ForEach(dataCell =>
                    {
                        tableRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(dataCell)))));
                    });
                    table.AppendChild(tableRow);
                });

                Body.AppendChild(table);
            }

            public void Build()
            {
                _document.Save();
                _document.Close();
            }
        }

        static void Main(string[] args)
        {
            var builder = new DocumentBuilder(args[0]);

            var table = new Table();
            var tableProperties = new TableProperties
            {
                TableBorders = new TableBorders
                {
                    TopBorder = new TopBorder
                        {Color = "000000", Val = BorderValues.Single, Size = 8, Space = 0},
                    RightBorder = new RightBorder
                        {Color = "000000", Val = BorderValues.Single, Size = 8, Space = 0},
                    BottomBorder = new BottomBorder
                        {Color = "000000", Val = BorderValues.Single, Size = 8, Space = 0},
                    LeftBorder = new LeftBorder
                        {Color = "000000", Val = BorderValues.Single, Size = 8, Space = 0},
                    InsideVerticalBorder = new InsideVerticalBorder
                        {Color = "000000", Val = BorderValues.Single, Size = 8, Space = 0},
                    InsideHorizontalBorder = new InsideHorizontalBorder
                        {Color = "000000", Val = BorderValues.Single, Size = 8, Space = 0},
                }
            };

            table.AppendChild(tableProperties);

            var tableRow = new TableRow();

            tableRow.AppendChild(
                new TableCell(new Paragraph(new Run(new Text("Recommendation:"))
                    {RunProperties = new RunProperties(new Bold())}))
                {
                    TableCellProperties = new TableCellProperties
                    {
                        VerticalMerge = new VerticalMerge {Val = MergedCellValues.Restart}
                    }
                });
            tableRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(
                "A bunch of example text etc etc etc and some more text to make the box quite large and stuff"))))
            {
                TableCellProperties = new TableCellProperties
                    {VerticalMerge = new VerticalMerge {Val = MergedCellValues.Restart}}
            });

            tableRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("Date:"))
                {RunProperties = new RunProperties(new Bold())})));
            tableRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("09/01/2020")))));

            var secondRow = new TableRow();
            secondRow.AppendChild(new TableCell(new Paragraph(new Run()))
            {
                TableCellProperties = new TableCellProperties
                    {VerticalMerge = new VerticalMerge {Val = MergedCellValues.Continue}}
            });
            secondRow.AppendChild(new TableCell(new Paragraph(new Run()))
            {
                TableCellProperties = new TableCellProperties
                    {VerticalMerge = new VerticalMerge {Val = MergedCellValues.Continue}}
            });
            secondRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("Author:"))
                {RunProperties = new RunProperties(new Bold())})));
            secondRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("Meow Meowington")))));

            var thirdRow = new TableRow();
            thirdRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("Is AO Required?"))
                {RunProperties = new RunProperties(new Bold())})));
            thirdRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("No")))));
            thirdRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("Cleared By:"))
                {RunProperties = new RunProperties(new Bold())})));
            thirdRow.AppendChild(new TableCell(new Paragraph(new Run(new Text("Barks Barkington")))));

            table.AppendChild(tableRow);
            table.AppendChild(secondRow);
            table.AppendChild(thirdRow);
            builder.Body.AppendChild(table);

            builder.Build();
            Console.WriteLine("Hello World!");
        }
    }
}