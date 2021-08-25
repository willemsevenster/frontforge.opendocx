using DocumentFormat.OpenXml;

namespace Frontforge.OpenDocx.Core.Models {
    public abstract class ContentElement
        : IContent
    {
        internal abstract OpenXmlElement Render(int index, bool isFirst, bool isLast);
    }
}