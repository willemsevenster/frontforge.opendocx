using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.ModelConfiguration;

namespace Frontforge.OpenDocx.Core.Models
{
    public class Section
    {
        private readonly SectionConfig _config;

        internal Body Render(int index, bool isFirst, bool isLast, MainDocumentPart mainPart)
        {
            var result = new Body(_config.SectionProperties(mainPart));

            foreach (var element in _config.Contents.Where(x => x != null).AsIndexed())
            {
                result.AppendChild(element.Value.Render(element.Index, element.IsFirst, element.IsLast));
            }

            return result;
        }

        internal Section(SectionConfig config)
        {
            _config = config;
        }
    }
}