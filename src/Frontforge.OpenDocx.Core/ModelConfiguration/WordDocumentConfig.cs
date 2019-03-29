using System.Collections.Generic;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.ModelConfiguration
{
    internal class WordDocumentConfig
    {
        #region properties and indexers

        #region properties

        public List<Section> Sections { get; } = new List<Section>();

        #endregion

        #endregion
    }
}