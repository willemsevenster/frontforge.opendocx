using DocumentFormat.OpenXml;
using Frontforge.OpenDocx.Core.ModelConfiguration;

namespace Frontforge.OpenDocx.Core.Models
{
    public class Paragraph
        : ContentElement
    {
        private readonly ParagraphConfig _config;

        internal Paragraph(ParagraphConfig config)
        {
            _config = config;
        }

        internal override OpenXmlElement Render(int index, bool isFirst, bool isLast)
        {
            var result = new DocumentFormat.OpenXml.Wordprocessing.Paragraph
            {
                ParagraphProperties = _config.GetParagraphProperties()
            };

            foreach (var run in _config.GetRuns())
            {
                result.AppendChild(run);
            }

            return result;
        }
    }
}
