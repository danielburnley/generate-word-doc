using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace Example3
{
    public class E3Text
    {
        public string Value { get; set; }
        public bool Bold { get; set; }
    }

    public class ExampleThreeBuilder
    {
        private readonly WordprocessingDocument _document;
        private readonly Body _body;

        public ExampleThreeBuilder(string filePath)
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

        public void AddParagraph(Action<ParagraphBuilder> func)
        {
            var paragraph = new Paragraph();
            var builder = new ParagraphBuilder(paragraph);
            func(builder);
            _body.AppendChild(paragraph);
        }

        public class TableBuilder
        {
            private readonly OpenXmlElement _parent;

            public TableBuilder(OpenXmlElement parent)
            {
                _parent = parent;
            }
        }

        public class TableRowBuilder
        {
            private readonly OpenXmlElement _parent;

            public TableRowBuilder(OpenXmlElement parent)
            {
                _parent = parent;
            }

            public void AddCell(string text)
            {
            }

            public void AddCells(string[] text)
            {
            }

            public void AddCell(E3Text text)
            {
            }

            public void AddCells(E3Text[] text)
            {
            }
        }

        public class ParagraphBuilder
        {
            private readonly OpenXmlElement _parent;

            public ParagraphBuilder(OpenXmlElement parent)
            {
                _parent = parent;
            }

            public class ParagraphTextFormattingOptions
            {
                public bool Bold { get; set; }
            }

            public void AddText(string text)
            {
                var run = new Run();
                var element = new Text(text)
                {
                    Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve)
                };

                run.AppendChild(element);
                _parent.AppendChild(run);
            }

            public void AddText(E3Text text)
            {
                var run = new Run();
                var element = new Text(text.Value)
                {
                    Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve)
                };

                if (text.Bold)
                {
                    run.RunProperties = new RunProperties
                    {
                        Bold = new Bold()
                    };
                }

                run.AppendChild(element);
                _parent.AppendChild(run);
            }
        }

        public void Build()
        {
            _document.Save();
            _document.Close();
        }
    }
}