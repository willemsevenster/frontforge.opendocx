using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.ModelConfiguration;
using Frontforge.OpenDocx.Core.Models;
using Paragraph = Frontforge.OpenDocx.Core.Models.Paragraph;

namespace Frontforge.OpenDocx.Core.Builders
{
    public class ParagraphBuilder
    {
        #region instance fields

        private readonly ParagraphConfig _config = new ParagraphConfig();

        #endregion

        #region implementation

        public static implicit operator Paragraph(ParagraphBuilder builder)
        {
            return new Paragraph(builder._config);
        }

        public ParagraphBuilder Add(string text)
        {
            return this.Chain(p => p._config.Contents.Add(new TextContent(text)));
        }

        public ParagraphBuilder Add(TextContent content)
        {
            return this.Chain(p => p._config.Contents.Add(content));
        }

        public ParagraphBuilder AddCheckbox(string label, bool isChecked = false)
        {
            var cb = (CheckboxControl) new CheckboxBuilder().Label(label).Check(isChecked);
            return this.Chain(p => p._config.Contents.Add(cb));
        }

        public ParagraphBuilder Alignment(HorizontalAlignment alignment)
        {
            return this.Chain(p => p._config.HorizontalAlignment = alignment);
        }

        public ParagraphBuilder AlignLeft()
        {
            return this.Chain(p => p._config.HorizontalAlignment = HorizontalAlignment.Left);
        }

        public ParagraphBuilder AlignRight()
        {
            return this.Chain(p => p._config.HorizontalAlignment = HorizontalAlignment.Right);
        }

        public ParagraphBuilder AlignCenter()
        {
            return this.Chain(p => p._config.HorizontalAlignment = HorizontalAlignment.Center);
        }

        public ParagraphBuilder AlignJustify()
        {
            return this.Chain(p => p._config.HorizontalAlignment = HorizontalAlignment.Justified);
        }

        public ParagraphBuilder Bold(bool bold = true)
        {
            return this.Chain(p => p._config.Bold = bold);
        }

        public ParagraphBuilder FontSize(Unit fontSize)
        {
            return this.Chain(p => p._config.FontSize = fontSize);
        }

        public ParagraphBuilder Bullet()
        {
            return this.Chain(p => p._config.Bullets = true);
        }

        #endregion

        #region constructors

        internal ParagraphBuilder()
        {
        }

        #endregion

        #region members

        public ParagraphBuilder SpacingBefore(Unit value)
        {
            return this.Chain(p => p._config.SpacingBefore = value);
        }

        public ParagraphBuilder SpacingAfter(Unit value)
        {
            return this.Chain(p => p._config.SpacingAfter = value);
        }

        public ParagraphBuilder TopBorder(uint size = 1U, BorderValues lineStyle = BorderValues.Single)
        {
            return this.Chain(p => p._config.Borders.TopBorder = new TopBorder
            {
                Val = lineStyle,
                Size = size
            });
        }

        public ParagraphBuilder BottomBorder(uint size = 1U, BorderValues lineStyle = BorderValues.Single)
        {
            return this.Chain(p => p._config.Borders.BottomBorder = new BottomBorder
            {
                Val = lineStyle,
                Size = size
            });
        }

        #endregion
    }
}