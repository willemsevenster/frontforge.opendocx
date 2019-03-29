using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.ModelConfiguration;

namespace Frontforge.OpenDocx.Core.Models
{
    public class TextContent
        : IContent
    {
        #region instance fields

        private readonly TextContentConfig _config;

        #endregion

        #region implementation

        //public static implicit operator TextContent(string text)
        //{
        //    return new TextContent {Value = text};
        //}

        public static implicit operator string(TextContent content)
        {
            return content?._config?.Value;
        }

        public Run GetRun(RunProperties runProperties)
        {
            var run = new Run {RunProperties = runProperties.CloneNode()};

            if (_config.Bold.HasValue)
            {
                run.RunProperties.Bold = new Bold {Val = OnOffValue.FromBoolean(_config.Bold.Value)};
            }

            if (_config.Italic.HasValue)
            {
                run.RunProperties.Italic = new Italic {Val = OnOffValue.FromBoolean(_config.Italic.Value)};
            }

            if (_config.Underline.HasValue)
            {
                run.RunProperties.Underline = new Underline {Val = _config.Underline};
            }

            if (!string.IsNullOrEmpty(this))
            {
                var lines = ((string) this).Split(new[] {"\r\n", "\n"}, StringSplitOptions.None).AsIndexed();

                foreach (var line in lines)
                {
                    if (!line.IsFirst)
                    {
                        run.AppendChild(new Break());
                    }

                    if (!string.IsNullOrEmpty(line.Value))
                    {
                        run.AppendChild(new Text(line) {Space = SpaceProcessingModeValues.Preserve});
                    }
                }
            }

            return run;
        }

        #endregion

        #region constructors

        internal TextContent(TextContentConfig config)
        {
            _config = config ?? new TextContentConfig();
        }

        internal TextContent(string text)
        {
            _config = new TextContentConfig {Value = text};
        }

        #endregion

        #region properties

        #endregion
    }
}