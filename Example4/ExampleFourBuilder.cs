using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Example4
{
    public class ExampleFourBuilder
    {
        private readonly WordprocessingDocument _document;
        private readonly Body _body;
        private readonly Example4DocumentBuilder _example4DocumentBuilder;

        public ExampleFourBuilder(string filePath, Example4DocumentBuilder example4DocumentBuilder)
        {
            _example4DocumentBuilder = example4DocumentBuilder;
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
    }
}