using System;
using System.Collections.Generic;
using Example1.DocumentBuilder;
using Example3;
using Example4;

namespace GenerateWordDoc
{
    static class Program
    {
        static void Main(string[] args)
        {
            ExampleOne(args);
            // ExampleThree(args);
        }


        private static void ExampleThree(IReadOnlyList<string> args)
        {
            var builder = new ExampleThreeBuilder(args[0]);
            builder.AddParagraph(pBuilder =>
            {
                pBuilder.AddText("Meow ");
                pBuilder.AddText(new E3Text {Value = "Meow", Bold = true});
                pBuilder.AddText(" Meow");
            });
            builder.Build();
            Console.WriteLine($"Generated document at: {args[0]}");
        }

        private static void ExampleOne(IReadOnlyList<string> args)
        {
            var builder = new ExampleOneBuilder($"{args[0]}");

            builder.AddHeading("Meow", ExampleOneBuilder.HeadingOptions.Heading1);
            builder.AddHeading("Sub meow", ExampleOneBuilder.HeadingOptions.Heading2);
            builder.AddParagraphToDocument(
                "Here is an introductory paragraph, it does some stuff.\r\nIt even has a new line in it, which as it turns out needs some manual poking");
            builder.AddLineBreak();

            var rows = new List<List<TableBuilder.BuilderTableCell>>
            {
                new List<TableBuilder.BuilderTableCell>
                {
                    new TableBuilder.BuilderTableCell
                        {Text = "Recommendation", Options = new TableBuilder.BuilderTableCellOptions {Bold = true}},
                    new TableBuilder.BuilderTableCell
                        {Text = "A bunch of example text etc etc with some more to pad the text out"},
                    new TableBuilder.BuilderTableCell
                        {Text = "Date:", Options = new TableBuilder.BuilderTableCellOptions {Bold = true}},
                    new TableBuilder.BuilderTableCell {Text = "01/09/2020"}
                },
                new List<TableBuilder.BuilderTableCell>
                {
                    new TableBuilder.BuilderTableCell
                        {Options = new TableBuilder.BuilderTableCellOptions {MergeAbove = true}},
                    new TableBuilder.BuilderTableCell
                        {Options = new TableBuilder.BuilderTableCellOptions {MergeAbove = true}},
                    new TableBuilder.BuilderTableCell
                        {Text = "Author:", Options = new TableBuilder.BuilderTableCellOptions {Bold = true}},
                    new TableBuilder.BuilderTableCell {Text = "Meow Meowington"}
                },
                new List<TableBuilder.BuilderTableCell>
                {
                    new TableBuilder.BuilderTableCell
                        {Text = "Is AO Required?", Options = new TableBuilder.BuilderTableCellOptions {Bold = true}},
                    new TableBuilder.BuilderTableCell {Text = "No"},
                    new TableBuilder.BuilderTableCell
                        {Text = "Cleared by:", Options = new TableBuilder.BuilderTableCellOptions {Bold = true}},
                    new TableBuilder.BuilderTableCell {Text = "Barks Barkington"}
                }
            };

            builder.AddTable(rows);
            builder.AddLineBreak();
            var listItems = new List<string> {"First line in a list", "then a second", "And then a third"};
            builder.AddBulletListToDocument(listItems);
            builder.Build();
            Console.WriteLine($"Generated document at: {args[0]}");
        }
        
        private static void ExampleFour(string[] args)
        {
            var builder = new Example4DocumentBuilder(args[0])
            {
                Children = new IDocumentElement[]
                {
                    new ParagraphBuilder
                    {
                        Children = new IDocumentElement[]
                        {
                            new TextBuilder {Val = "Meow "},
                            new TextBuilder {Val = "Woof", Bold = true},
                            new TextBuilder {Val = " Meow"}
                        }
                    },
                    new TableBuilderE4
                    {
                        Children = new IDocumentElement[]
                        {
                            new TableRowBuilder()
                            {
                                Children = new IDocumentElement[]
                                {
                                    new TableCellBuilder()
                                    {
                                        Children = new IDocumentElement[]
                                        {
                                            new ParagraphBuilder()
                                            {
                                                Children = new IDocumentElement[]
                                                {
                                                    new TextBuilder {Val = "Meow "},
                                                    new TextBuilder {Val = "Woof", Bold = true},
                                                    new TextBuilder {Val = " Meow"}
                                                }
                                            }
                                        }
                                    },
                                    new TableCellBuilder()
                                    {
                                        Children = new IDocumentElement[]
                                        {
                                            new ParagraphBuilder()
                                            {
                                                Children = new IDocumentElement[]
                                                {
                                                    new TextBuilder {Val = "Meow "},
                                                    new TextBuilder {Val = "Woof", Bold = true},
                                                    new TextBuilder {Val = " Meow"}
                                                }
                                            }
                                        }
                                    }
                                }
                            },
                            new TableRowBuilder()
                            {
                                Children = new IDocumentElement[]
                                {
                                    new TableCellBuilder()
                                    {
                                        Children = new IDocumentElement[]
                                        {
                                            new ParagraphBuilder()
                                            {
                                                Children = new IDocumentElement[]
                                                {
                                                    new TextBuilder {Val = "Meow "},
                                                    new TextBuilder {Val = "Woof", Bold = true},
                                                    new TextBuilder {Val = " Meow"}
                                                }
                                            }
                                        }
                                    },
                                    new TableCellBuilder()
                                    {
                                        Children = new IDocumentElement[]
                                        {
                                            new ParagraphBuilder()
                                            {
                                                Children = new IDocumentElement[]
                                                {
                                                    new TextBuilder {Val = "Meow "},
                                                    new TextBuilder {Val = "Woof", Bold = true},
                                                    new TextBuilder {Val = " Meow"}
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };
            builder.Build();
        }
    }
}