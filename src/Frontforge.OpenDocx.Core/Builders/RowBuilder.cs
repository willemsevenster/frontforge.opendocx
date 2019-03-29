using System.Linq;
using Frontforge.OpenDocx.Core.ModelConfiguration;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.Builders
{
    public class RowBuilder
    {
        #region instance fields

        private readonly RowConfig _config = new RowConfig();

        #endregion

        #region implementation

        public RowBuilder EnsureColumns(int columns)
        {
            return this.Chain(p =>
            {
                if (p._config.Cells.Count >= columns)
                {
                    return;
                }

                foreach (var col in Enumerable.Range(1, columns - p._config.Cells.Count))
                {
                    p._config.Cells.Add(CellBuilder.EmptyCell());
                }
            });
        }

        #endregion

        #region constructors

        internal RowBuilder()
        {
        }

        #endregion

        #region members

        public static implicit operator Row(RowBuilder builder)
        {
            return new Row(builder._config);
        }

        public RowBuilder Add(params Cell[] cells)
        {
            // null cells are translated to empty cells.
            return this.Chain(p =>
            {
                if (cells != null)
                {
                    p._config.Cells.AddRange(cells.Select(c => c ?? new CellBuilder()));
                }
            });
        }

        #endregion
    }
}