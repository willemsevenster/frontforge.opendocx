using DocumentFormat.OpenXml;
using Frontforge.OpenDocx.Core.ModelConfiguration;

namespace Frontforge.OpenDocx.Core.Models
{
    public class Paragraph
        : ContentElement
    {
        #region instance fields

        private readonly ParagraphConfig _config;

        #endregion

        #region implementation

        #region members

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

        #endregion

        #endregion

        #region constructors

        internal Paragraph(ParagraphConfig config)
        {
            _config = config;
        }

        #endregion
    }
}