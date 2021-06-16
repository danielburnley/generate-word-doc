using System;
using System.Collections.Generic;
using Example1.DocumentBuilder;
using static Example1.DocumentBuilder.TableBuilder;

namespace GenerateWordDoc
{
    static class Program
    {
        static void Main(string[] args)
        {
            var builder = new ExampleOneBuilder(args[0]);

            builder.AddHeading("Meow", ExampleOneBuilder.HeadingOptions.Heading1);
            builder.AddHeading("Sub meow", ExampleOneBuilder.HeadingOptions.Heading2);
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
            var listItems = new List<string> {"First line in a list", "then a second", "And then a third"};
            builder.AddBulletListToDocument(listItems);
            builder.Build();
            Console.WriteLine($"Generated document at: {args[0]}");
        }
    }
}