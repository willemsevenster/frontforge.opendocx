using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.ModelConfiguration
{
    internal class RowConfig
    {
        #region properties and indexers

        #region properties

        public List<Cell> Cells { get; } = new List<Cell>();

        #endregion

        #endregion

        #region implementation

        #region members

        public TableRowProperties GetRowProperties()
        {
            var result = new TableRowProperties();

            return result;
        }

        #endregion

        #endregion
    }
}