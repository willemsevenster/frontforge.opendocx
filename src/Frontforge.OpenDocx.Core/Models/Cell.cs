using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.ModelConfiguration;

namespace Frontforge.OpenDocx.Core.Models
{
    public class Cell
        : ContentElement
    {
        private readonly CellConfig _config;

        internal override OpenXmlElement Render(int index, bool isFirst, bool isLast)
        {
            var result = new TableCell(_config.CellProperties());

            var elements = _config.Contents.Where(x => x != null).AsIndexed().ToList();

            if (!elements.Any()) result.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());

            foreach (var element in elements)
            {
                result.AppendChild(element.Value.Render(element.Index, element.IsFirst, element.IsLast));
            }

            return result;
        }

        internal Cell(CellConfig config)
        {
            _config = config;
        }
    }
}