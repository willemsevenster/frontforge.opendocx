using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.ModelConfiguration;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.Builders
{
    public class TextContentBuilder
    {
        private readonly TextContentConfig _config = new TextContentConfig();

        internal TextContentBuilder(string text)
        {
            _config.Value = text;
        }

        public static implicit operator TextContent(TextContentBuilder builder)
        {
            return new TextContent(builder._config);
        }

        public TextContentBuilder FontSize(Unit fontSize)
        {
            return this.Chain(p => p._config.FontSize = fontSize);
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
    }
}
