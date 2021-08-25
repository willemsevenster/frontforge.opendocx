using DocumentFormat.OpenXml;
using Frontforge.OpenDocx.Core.ModelConfiguration;

namespace Frontforge.OpenDocx.Core.Models
{
    public class Table
        : ContentElement
    {
        private readonly TableConfig _config;

        internal override OpenXmlElement Render(int index, bool isFirst, bool isLast)
        {
            var result = new DocumentFormat.OpenXml.Wordprocessing.Table(_config.GetTableProperties());

            foreach (var row in _config.Rows.AsIndexed())
            {
                result.AppendChild(row.Value.Render(row.Index, row.IsFirst, row.IsLast));
            }

            return result;
        }

        internal Table(TableConfig config)
        {
            _config = config;
        }
    }
}