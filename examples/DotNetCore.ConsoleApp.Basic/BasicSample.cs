using Frontforge.OpenDocx.Core;
using Frontforge.OpenDocx.Core.Models;

namespace DotNetCore.ConsoleApp.Basic
{
    internal class BasicSample
        : WordDocument
    {
        #region implementation

        public static BasicSample Create()
        {
            return new BasicSample().BuildDoc();
        }

        private BasicSample BuildDoc()
        {
            // create a new section and set section properties
            var section = Section()
                .PageSize(PageSize.A4)
                .PageMargins(PredefinedPageMargins.Narrow);

            // add a paragraph
            var par = Par("This is a simple paragraph that is bold and center aligned with a " +
                          "16pt font size and a 6pt spacing before the paragraph.",
                    HorizontalAlignment.Center)
                .SpacingBefore(6)
                .Bold()
                .FontSize(16);

            section.Add(par);

            // add a table
            var tbl = Table()
                .Width(new Unit(100, UnitType.pct)) // 100% width
                .TopBorder()
                .BottomBorder();

            tbl.Add(
                Row(
                    Cell(Par("Cell 0, 0").Bold()).Width(30, UnitType.pct),
                    Cell(Par("Cell 0, 1"))
                ),
                Row(
                    Cell(Par("Cell 1, 0").Bold()),
                    Cell(Par("Cell 1, 1"))
                )
            );

            // add the table to the section
            section.Add(tbl);

            // add the section to the document
            AddSection(section);

            return this;
        }

        #endregion
    }
}