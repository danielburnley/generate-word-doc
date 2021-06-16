using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Example1.DocumentBuilder
{
    public static class DocumentBuilderHelpers
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
}