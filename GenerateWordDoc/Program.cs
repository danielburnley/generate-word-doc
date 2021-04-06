using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GenerateWordDoc
{
    class Program
    {
        private class BuilderTableCell
        {
            public string Text = "";
            public BuilderTableCellOptions Options;
        }

        private class BuilderTableCellOptions
        {
            public bool Bold;
            public bool MergeAbove;
        }

        private static class DocumentBuilderHelpers
        {
            public static void AddTextToElement(OpenXmlElement element, string text)
            {
                var splitText = text.Split(new[] {"\r\n", "\n", "\r"}, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in splitText)
                {
                    element.AppendChild(new Text(line));
                    if (line != splitText.Last())
                    {
                        element.AppendChild(new Break());
                    }
                }
            }
        }

        private class TableBuilder
        {
            private readonly Table _table;

            public TableBuilder()
            {
                _table = new Table();

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

                _table.AppendChild(tableProperties);
            }

            public void AddRowToTable(List<BuilderTableCell> rowCells)
            {
                var tableRow = new TableRow();
                rowCells.ForEach(dataCell =>
                {
                    var tableCell = new TableCell();
                    var tableCellProperties = new TableCellProperties
                    {
                        VerticalMerge = new VerticalMerge
                        {
                            Val = MergedCellValues.Restart
                        }
                    };

                    if (dataCell.Options != null && dataCell.Options.MergeAbove)
                    {
                        tableCellProperties.VerticalMerge.Val = MergedCellValues.Continue;
                    }

                    tableCell.TableCellProperties = tableCellProperties;

                    var paragraph = new Paragraph();
                    var run = new Run();
                    DocumentBuilderHelpers.AddTextToElement(run, dataCell.Text);

                    if (dataCell.Options != null && dataCell.Options.Bold)
                    {
                        run.RunProperties = new RunProperties(new Bold());
                    }

                    paragraph.AppendChild(run);
                    tableCell.AppendChild(paragraph);
                    tableRow.AppendChild(tableCell);
                });
                _table.AppendChild(tableRow);
            }

            public void Build(OpenXmlElement parentElement)
            {
                parentElement.AppendChild(_table);
            }
        }

        private class DocumentBuilder
        {
            private readonly WordprocessingDocument _document;
            private readonly Body _body;

            public DocumentBuilder(string filePath)
            {
                _document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
                var mainPart = _document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                _body = mainPart.Document.Body;
            }

            public void AddLineBreak()
            {
                _body.AppendChild(new Paragraph(new Break()));
            }

            public void AddParagraphToDocument(string text)
            {
                var para = new Paragraph();
                var run = new Run();
                DocumentBuilderHelpers.AddTextToElement(run, text);
                para.AppendChild(run);
                _body.AppendChild(para);
            }

            public void AddTable(List<List<BuilderTableCell>> data)
            {
                var tableBuilder = new TableBuilder();

                data.ForEach(dataRow => { tableBuilder.AddRowToTable(dataRow); });

                tableBuilder.Build(_body);
            }

            public void Build()
            {
                _document.Save();
                _document.Close();
            }

            public enum HeadingOptions
            {
                Heading1,
                Heading2
            }

            private static readonly Dictionary<HeadingOptions, string> HeadingSizes =
                new Dictionary<HeadingOptions, string>
                {
                    {HeadingOptions.Heading1, "80"}
                };

            public void AddHeading(string text, HeadingOptions level)
            {
                _body.AppendChild(new Paragraph(new Run(new Text(text))
                {
                    RunProperties = new RunProperties
                    {
                        FontSize = new FontSize {Val = HeadingSizes[level]}
                    }
                }));
            }
        }

        static void Main(string[] args)
        {
            var builder = new DocumentBuilder(args[0]);

            builder.AddHeading("Meow", DocumentBuilder.HeadingOptions.Heading1);
            builder.AddParagraphToDocument(
                "Here is an introductory paragraph, it does some stuff.\r\nIt even has a new line in it, which as it turns out needs some manual poking");
            builder.AddLineBreak();

            var rows = new List<List<BuilderTableCell>>
            {
                new List<BuilderTableCell>
                {
                    new BuilderTableCell {Text = "Recommendation", Options = new BuilderTableCellOptions {Bold = true}},
                    new BuilderTableCell {Text = "A bunch of example text etc etc with some more to pad the text out"},
                    new BuilderTableCell {Text = "Date:", Options = new BuilderTableCellOptions {Bold = true}},
                    new BuilderTableCell {Text = "01/09/2020"}
                },
                new List<BuilderTableCell>
                {
                    new BuilderTableCell {Options = new BuilderTableCellOptions {MergeAbove = true}},
                    new BuilderTableCell {Options = new BuilderTableCellOptions {MergeAbove = true}},
                    new BuilderTableCell {Text = "Author:", Options = new BuilderTableCellOptions {Bold = true}},
                    new BuilderTableCell {Text = "Meow Meowington"}
                },
                new List<BuilderTableCell>
                {
                    new BuilderTableCell
                        {Text = "Is AO Required?", Options = new BuilderTableCellOptions {Bold = true}},
                    new BuilderTableCell {Text = "No"},
                    new BuilderTableCell {Text = "Cleared by:", Options = new BuilderTableCellOptions {Bold = true}},
                    new BuilderTableCell {Text = "Barks Barkington"}
                }
            };

            builder.AddTable(rows);
            builder.AddLineBreak();
            builder.AddTable(rows);
            builder.Build();
            Console.WriteLine($"Generated document at: {args[0]}");
        }
    }
}