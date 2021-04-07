using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GenerateWordDoc.DocumentBuilder
{
    public class Builder
    {
        private readonly WordprocessingDocument _document;
        private readonly Body _body;

        public Builder(string filePath)
        {
            _document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
            var mainPart = _document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            _body = mainPart.Document.Body;
            AddNumberingDefinitions();
        }

        private void AddNumberingDefinitions()
        {
            var numberingDefinitionsPart =
                _document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>("unorderedList");

            var element = new Numbering(
                new AbstractNum(
                    new Level(
                        new NumberingFormat {Val = NumberFormatValues.Bullet},
                        new LevelJustification {Val = LevelJustificationValues.Left},
                        new LevelText {Val = "●"},
                        new RunProperties
                        {
                            RunFonts = new RunFonts
                            {
                                Ascii = "Symbol",
                                HighAnsi = "Symbol",
                                ComplexScript = "Symbol",
                                Hint = FontTypeHintValues.Default
                            }
                        }
                    ) {LevelIndex = 0}
                ) {AbstractNumberId = 1},
                new NumberingInstance(new AbstractNumId {Val = 1}) {NumberID = 1}
            );

            element.Save(numberingDefinitionsPart);
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

        public void AddBulletListToDocument(List<string> listItems)
        {
            listItems.ForEach(item => _body.AppendChild(
                new Paragraph(
                    new ParagraphProperties(
                        new NumberingProperties(
                            new NumberingLevelReference {Val = 0},
                            new NumberingId {Val = 1})),
                    new Run(
                        new RunProperties(),
                        new Text(item) {Space = SpaceProcessingModeValues.Preserve}))
            ));
        }

        public void AddTable(List<List<TableBuilder.BuilderTableCell>> data)
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
                {HeadingOptions.Heading1, "80"},
                {HeadingOptions.Heading2, "60"}
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
}