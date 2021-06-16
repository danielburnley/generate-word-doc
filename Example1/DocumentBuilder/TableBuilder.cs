using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Example1.DocumentBuilder
{
    public class TableBuilder
    {
        public class BuilderTableCell
        {
            public string Text = "";
            public BuilderTableCellOptions Options;
        }

        public class BuilderTableCellOptions
        {
            public bool Bold;
            public bool MergeAbove;
        }

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
}