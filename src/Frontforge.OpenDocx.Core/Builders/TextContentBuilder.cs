using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.ModelConfiguration;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.Builders
{
    public class TextContentBuilder
    {
        #region instance fields

        private readonly TextContentConfig _config = new TextContentConfig();

        #endregion

        #region implementation

        public static implicit operator TextContent(TextContentBuilder builder)
        {
            return new TextContent(builder._config);
        }

        public TextContentBuilder Bold(bool? bold = true)
        {
            return this.Chain(p => p._config.Bold = bold);
        }

        public TextContentBuilder Italic(bool? italic = true)
        {
            return this.Chain(p => p._config.Italic = italic);
        }

        public TextContentBuilder Underline(UnderlineValues? line = UnderlineValues.Single)
        {
            return this.Chain(p => p._config.Underline = line);
        }

        #endregion

        #region constructors

        internal TextContentBuilder(string text)
        {
            _config.Value = text;
        }

        #endregion
    }
}