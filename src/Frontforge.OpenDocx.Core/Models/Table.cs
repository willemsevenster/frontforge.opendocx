using DocumentFormat.OpenXml;
using Frontforge.OpenDocx.Core.ModelConfiguration;

namespace Frontforge.OpenDocx.Core.Models
{
    public class Table
        : ContentElement
    {
        #region instance fields

        private readonly TableConfig _config;

        #endregion

        #region implementation

        #region members

        internal override OpenXmlElement Render(int index, bool isFirst, bool isLast)
        {
            var result = new DocumentFormat.OpenXml.Wordprocessing.Table(_config.GetTableProperties());

            foreach (var row in _config.Rows.AsIndexed())
            {
                result.AppendChild(row.Value.Render(row.Index, row.IsFirst, row.IsLast));
            }

            return result;
        }

        #endregion

        #endregion

        #region constructors

        internal Table(TableConfig config)
        {
            _config = config;
        }

        #endregion
    }
}