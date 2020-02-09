using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.ModelConfiguration;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.Builders
{
    public class CellBuilder
    {
        #region instance fields

        private readonly CellConfig _config = new CellConfig();

        #endregion

        #region constructors

        internal CellBuilder() { }

        #endregion

        #region implementation

        internal static CellBuilder EmptyCell()
        {
            return new CellBuilder().Add(new ParagraphBuilder().SpacingBefore(0).SpacingAfter(0));
        }

        #endregion

        #region members

        public static implicit operator Cell(CellBuilder builder)
        {
            return new Cell(builder._config);
        }

        public CellBuilder Add(params ContentElement[] contents)
        {
            return this.Chain(p => p._config.Contents.AddRange(contents.Where(x => x != null)));
        }

        public CellBuilder Width(double value, UnitType type = UnitType.pct)
        {
            return this.Chain(p => p._config.Width = new Unit(value, type));
        }

        public CellBuilder BgColor(string color, ShadingPatternValues pattern = ShadingPatternValues.Percent50)
        {
            return this.Chain(p =>
            {
                p._config.BackgroundColor.Color = color;
                p._config.BackgroundColor.Fill = color;
                p._config.BackgroundColor.Val = pattern;
            });
        }

        public CellBuilder TopBorder(uint size = 1U, BorderValues lineStyle = BorderValues.Single, string color = null)
        {
            return this.Chain(p => p._config.Borders.TopBorder = new TopBorder
            {
                Val = lineStyle,
                Size = size,
                Color = color
            });
        }

        public CellBuilder LeftBorder(uint size = 1U, BorderValues lineStyle = BorderValues.Single, string color = null)
        {
            return this.Chain(p => p._config.Borders.LeftBorder = new LeftBorder
            {
                Val = lineStyle,
                Size = size,
                Color = color
            });
        }

        public CellBuilder RightBorder(uint size = 1U, BorderValues lineStyle = BorderValues.Single, string color = null)
        {
            return this.Chain(p => p._config.Borders.RightBorder = new RightBorder
            {
                Val = lineStyle,
                Size = size,
                Color = color
            });
        }

        public CellBuilder BottomBorder(uint size = 1U, BorderValues lineStyle = BorderValues.Single, string color = null)
        {
            return this.Chain(p => p._config.Borders.BottomBorder = new BottomBorder
            {
                Val = lineStyle,
                Size = size,
                Color = color
            });
        }

        public CellBuilder AllBorders(uint size = 1U, BorderValues lineStyle = BorderValues.Single, string color = null)
        {
            return this.Chain(p => p.TopBorder(size, lineStyle, color)
                .LeftBorder(size, lineStyle, color)
                .RightBorder(size, lineStyle, color)
                .BottomBorder(size, lineStyle, color));
        }

        public CellBuilder VAlignTop()
        {
            return this.Chain(p => p._config.VerticalAlignment = TableVerticalAlignmentValues.Top);
        }

        public CellBuilder VAlignMiddle()
        {
            return this.Chain(p => p._config.VerticalAlignment = TableVerticalAlignmentValues.Center);
        }

        public CellBuilder Span(int colspan)
        {
            return this.Chain(p => p._config.ColSpan = colspan);
        }

        #endregion
    }
}
