﻿using DocumentFormat.OpenXml.Wordprocessing;

namespace Frontforge.OpenDocx.Core.ModelConfiguration
{
    internal class TextContentConfig
    {
        #region properties and indexers

        public bool? Bold { get; set; }

        public bool? Italic { get; set; }

        public UnderlineValues? Underline { get; set; }

        public string Value { get; set; }

        #endregion
    }
}