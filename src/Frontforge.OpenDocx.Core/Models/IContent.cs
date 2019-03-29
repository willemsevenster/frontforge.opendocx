using DocumentFormat.OpenXml;

namespace Frontforge.OpenDocx.Core.Models
{
    public interface IContent
    {
    }

    public abstract class ContentElement
        : IContent
    {
        #region implementation

        #region members

        internal abstract OpenXmlElement Render(int index, bool isFirst, bool isLast);

        #endregion

        #endregion
    }
}