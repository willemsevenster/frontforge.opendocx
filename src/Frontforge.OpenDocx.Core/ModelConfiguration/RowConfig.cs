using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.ModelConfiguration
{
    internal class RowConfig
    {
        public List<Cell> Cells { get; } = new List<Cell>();

        public TableRowProperties GetRowProperties()
        {
            var result = new TableRowProperties();

            return result;
        }
    }
}