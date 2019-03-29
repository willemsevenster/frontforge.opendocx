using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.ModelConfiguration
{
    internal class CellConfig
    {
        #region implementation

        #region members

        public TableCellProperties CellProperties()
        {
            var result = new TableCellProperties();

            if (NoWrap.HasValue)
            {
                result.NoWrap = new NoWrap {Val = NoWrap.Value ? OnOffOnlyValues.On : OnOffOnlyValues.Off};
            }

            if (Width != null)
            {
                result.TableCellWidth = Width.ToTableWidthType<TableCellWidth>();
            }

            result.TableCellBorders = Borders;

            result.TableCellVerticalAlignment = new TableCellVerticalAlignment
                {Val = VerticalAlignment};

            if (ColSpan > 1)
            {
                result.GridSpan = new GridSpan {Val = ColSpan};
            }

            return result;
        }

        #endregion

        #endregion

        #region properties

        public bool? NoWrap { get; set; }

        public List<ContentElement> Contents { get; } = new List<ContentElement>();

        public Unit Width { get; set; }

        public TableCellBorders Borders { get; } = new TableCellBorders();

        public int ColSpan { get; set; } = 1;

        public TableVerticalAlignmentValues VerticalAlignment { get; set; } = TableVerticalAlignmentValues.Bottom;

        #endregion
    }
}