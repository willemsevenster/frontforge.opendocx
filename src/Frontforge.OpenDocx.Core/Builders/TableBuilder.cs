using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.ModelConfiguration;
using Frontforge.OpenDocx.Core.Models;
using Table = Frontforge.OpenDocx.Core.Models.Table;

namespace Frontforge.OpenDocx.Core.Builders
{
    public class TableBuilder
    {
        #region instance fields

        private readonly TableConfig _config = new TableConfig();

        #endregion

        #region constructors

        internal TableBuilder()
        {
        }

        #endregion

        #region members

        public static implicit operator Table(TableBuilder builder)
        {
            return new Table(builder._config);
        }

        public TableBuilder Add(params Row[] rows)
        {
            return this.Chain(p =>
            {
                if (rows != null)
                {
                    p._config.Rows.AddRange(rows.Where(r => r != null));
                }
            });
        }

        public TableBuilder Width(Unit unit)
        {
            return this.Chain(p => p._config.Width = unit);
        }

        public TableBuilder CellMargins(Unit all)
        {
            return this.Chain(p =>
            {
                p._config.CellLeftMargin = all;
                p._config.CellRightMargin = all;
                p._config.CellTopMargin = all;
                p._config.CellBottomMargin = all;
            });
        }

        public TableBuilder CellMargins(Unit leftRight, Unit topBottom)
        {
            return this.Chain(p =>
            {
                p._config.CellLeftMargin = leftRight;
                p._config.CellRightMargin = leftRight;
                p._config.CellTopMargin = topBottom;
                p._config.CellBottomMargin = topBottom;
            });
        }

        public TableBuilder CellMargins(Unit left, Unit top, Unit bottom, Unit right)
        {
            return this.Chain(p =>
            {
                if (left != null)
                {
                    p._config.CellLeftMargin = left;
                }

                if (top != null)
                {
                    p._config.CellTopMargin = top;
                }

                if (bottom != null)
                {
                    p._config.CellBottomMargin = bottom;
                }

                if (right != null)
                {
                    p._config.CellRightMargin = right;
                }
            });
        }

        public TableBuilder CellSpacing(Unit value)
        {
            return this.Chain(p => p._config.CellSpacing = value);
        }


        public TableBuilder TopBorder(uint size = 1U, BorderValues lineStyle = BorderValues.Single)
        {
            return this.Chain(p => p._config.Borders.TopBorder = new TopBorder
            {
                Val = lineStyle,
                Size = size
            });
        }

        public TableBuilder BottomBorder(uint size = 1U, BorderValues lineStyle = BorderValues.Single)
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