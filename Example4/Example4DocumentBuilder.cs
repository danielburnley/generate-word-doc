using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Example4
{
    public class Example4DocumentBuilder
    {
        private readonly WordprocessingDocument _document;
        private readonly Body _body;
        public IDocumentElement[] Children { get; set; }

        public Example4DocumentBuilder(string filePath)
        {
            _document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
            _document.AddMainDocumentPart();
            _document.MainDocumentPart.Document = new Document(new Body());
            _body = _document.MainDocumentPart.Document.Body;
            SetCompatibilityMode();
        }

        private void SetCompatibilityMode()
        {
            var mainPart = _document.MainDocumentPart;
            var settingsPart = mainPart.DocumentSettingsPart;

            if (settingsPart != null) return;

            settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(
                new Compatibility(
                    new CompatibilitySetting
                    {
                        Name = new EnumValue<CompatSettingNameValues>
                            (CompatSettingNameValues.CompatibilityMode),
                        Val = new StringValue("15"),
                        Uri = new StringValue
                            ("http://schemas.microsoft.com/office/word")
                    }
                )
            );
            settingsPart.Settings.Save();
        }

        public void Build()
        {
            foreach (var child in Children)
            {
                _body.AppendChild(child.Build());
            }

            _document.Save();
            _document.Close();
        }
    }

    public interface IDocumentElement
    {
        public OpenXmlElement Build();
    }

    public class ParagraphBuilder : IDocumentElement
    {
        public IDocumentElement[] Children { get; set; }

        public OpenXmlElement Build()
        {
            var paragraph = new Paragraph();

            foreach (var child in Children)
            {
                var res = child.Build();
                paragraph.AppendChild(res);
            }

            return paragraph;
        }
    }

    public class TextBuilder : IDocumentElement
    {
        public string Val { get; set; }
        public bool Bold { get; set; }

        public OpenXmlElement Build()
        {
            var run = new Run();
            var element = new Text(Val)
            {
                Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve)
            };

            if (Bold)
            {
                run.RunProperties = new RunProperties
                {
                    Bold = new Bold()
                };
            }

            run.AppendChild(element);
            return run;
        }
    }

    public class TableBuilderE4 : IDocumentElement
    {
        public IDocumentElement[] Children { get; set; }

        private void AddTableProperties(Table table)
        {
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
                },
                TableCellMarginDefault = new TableCellMarginDefault
                {
                    TopMargin = new TopMargin {Width = "40"},
                    BottomMargin = new BottomMargin {Width = "40"},
                    TableCellRightMargin = new TableCellRightMargin {Width = 40},
                    TableCellLeftMargin = new TableCellLeftMargin {Width = 40}
                }
            };

            table.AppendChild(tableProperties);
        }

        public OpenXmlElement Build()
        {
            var table = new Table();
            AddTableProperties(table);

            foreach (var child in Children)
            {
                var res = child.Build();
                table.AppendChild(res);
            }

            return table;
        }
    }

    public class TableRowBuilder : IDocumentElement
    {
        public IDocumentElement[] Children { get; set; }

        public OpenXmlElement Build()
        {
            var tableRow = new TableRow();

            foreach (var child in Children)
            {
                tableRow.AppendChild(child.Build());
            }

            return tableRow;
        }
    }

    public class TableCellBuilder : IDocumentElement
    {
        public IDocumentElement[] Children { get; set; }

        public OpenXmlElement Build()
        {
            var tableCell = new TableCell();

            foreach (var child in Children)
            {
                tableCell.AppendChild(child.Build());
            }

            return tableCell;
        }
    }
}
// public void AddTableRow(List<string> cellData)
// {
//
//     var tableRow = new TableRow();
//
//     cellData.ForEach(cellText =>
//     {
//         var tableCell = new TableCell();
//
//         var paragraph = new Paragraph();
//         var run = new Run();
//         DocumentBuilderHelpers.AddTextToElement(run, cellText);
//
//         paragraph.AppendChild(run);
//         tableCell.AppendChild(paragraph);
//         tableRow.AppendChild(tableCell);
//     });
//
//     _table.AppendChild(tableRow);
// }