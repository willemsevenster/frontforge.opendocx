using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.Models;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;

namespace Frontforge.OpenDocx.Core.ModelConfiguration
{
    internal class TableConfig
    {
        #region implementation

        #region members

        public TableProperties GetTableProperties()
        {
            var result = new TableProperties();

            if (Width != null)
            {
                result.AppendChild(Width.ToTableWidthType<TableWidth>());
            }

            result.TableCellMarginDefault = new TableCellMarginDefault();

            if (CellLeftMargin != null)
            {
                result.TableCellMarginDefault.TableCellLeftMargin = CellLeftMargin.ToTableMargin<TableCellLeftMargin>();
            }

            if (CellTopMargin != null)
            {
                result.TableCellMarginDefault.TopMargin = CellTopMargin.ToTableWidthType<TopMargin>();
            }

            if (CellRightMargin != null)
            {
                result.TableCellMarginDefault.TableCellRightMargin =
                    CellRightMargin.ToTableMargin<TableCellRightMargin>();
            }

            if (CellBottomMargin != null)
            {
                result.TableCellMarginDefault.BottomMargin = CellBottomMargin.ToTableWidthType<BottomMargin>();
            }

            if (CellSpacing != null)
            {
                result.TableCellSpacing = new TableCellSpacing
                {
                    Width = CellSpacing
                };
            }

            result.TableBorders = Borders;

            return result;
        }

        #endregion

        #endregion

        #region properties

        public Unit Width { get; set; }

        public List<Row> Rows { get; } = new List<Row>();

        public Unit CellLeftMargin { get; set; }
        public Unit CellRightMargin { get; set; }
        public Unit CellTopMargin { get; set; }
        public Unit CellBottomMargin { get; set; }

        public Unit CellSpacing { get; set; }

        public TableBorders Borders { get; } = new TableBorders();

        #endregion
    }
}