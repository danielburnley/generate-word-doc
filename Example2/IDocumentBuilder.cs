using System.Collections.Generic;

namespace ExampleTwo
{
    public interface IDocumentBuilder
    {
        public void AddParagraph(string text);

        public void AddHeading(string text, DocumentHeadingBuilder.HeadingLevelOptions headingLevel);

        public void AddTable(List<List<string>> tableData);

        public void Build();
    }
}