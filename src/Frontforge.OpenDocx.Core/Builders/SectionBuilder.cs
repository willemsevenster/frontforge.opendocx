using System.Linq;
using Frontforge.OpenDocx.Core.ModelConfiguration;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.Builders
{
    public class SectionBuilder
    {
        #region instance fields

        private readonly SectionConfig _config = new SectionConfig();

        #endregion

        #region implementation

        public static implicit operator Section(SectionBuilder builder)
        {
            return new Section(builder._config);
        }

        public SectionBuilder Add(params ContentElement[] contents)
        {
            return this.Chain(p => p._config.Contents.AddRange(contents.Where(x => x != null)));
        }

        public SectionBuilder Orientation(PageOrientation orientation)
        {
            return this.Chain(p => p._config.Orientation = orientation);
        }

        public SectionBuilder PageSize(PageSize pageSize)
        {
            return this.Chain(p => p._config.PageSize = pageSize);
        }

        public SectionBuilder PageMargins(PredefinedPageMargins pageMargins)
        {
            return this.Chain(p => p._config.PageMargins = pageMargins);
        }

        public SectionBuilder AddHeader(params ContentElement[] contents)
        {
            return this.Chain(p => p._config.Header.AddRange(contents.Where(x => x != null)));
        }

        public SectionBuilder AddFooter(params ContentElement[] contents)
        {
            return this.Chain(p => p._config.Footer.AddRange(contents.Where(x => x != null)));
        }

        #endregion

        #region constructors

        internal SectionBuilder()
        {
        }

        #endregion
    }
}