using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.ModelConfiguration
{
    internal class WordDocumentConfig
    {
        public List<Section> Sections { get; } = new List<Section>();

        public List<ImageContentItem> ImageMedia { get; } = new List<ImageContentItem>();
    }
}