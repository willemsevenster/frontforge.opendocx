using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.ModelConfiguration
{
    internal class TextContentConfig
    {
        public bool? Bold { get; set; }

        public bool? Italic { get; set; }

        public UnderlineValues? Underline { get; set; }

        public string Value { get; set; }

        public Unit FontSize { get; set; }
    }
}