using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.ModelConfiguration;

namespace Frontforge.OpenDocx.Core.Models
{
    public class Row
        : ContentElement
    {
        #region instance fields

        private readonly RowConfig _config;

        #endregion

        #region implementation

        #region members

        internal override OpenXmlElement Render(int index, bool isFirst, bool isLast)
        {
            var result = new TableRow(_config.GetRowProperties());

            foreach (var cell in _config.Cells.AsIndexed())
            {
                result.AppendChild(cell.Value.Render(cell.Index, cell.IsFirst, cell.IsLast));
            }

            return result;
        }

        #endregion

        #endregion

        #region constructors

        internal Row(RowConfig config)
        {
            _config = config;
        }

        #endregion
    }
}