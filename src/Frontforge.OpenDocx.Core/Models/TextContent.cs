﻿using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.ModelConfiguration;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
namespace Frontforge.OpenDocx.Core.Models
{
    public class TextContent
        : IContent
    {
        private readonly TextContentConfig _config;

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

            if (_config.FontSize != null)
            {
                run.RunProperties.FontSize = new FontSize{Val = _config.FontSize};
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

            if (_config.Value == null)
            {
                run.AppendChild(new Break());
            }

            return run;
        }

        internal TextContent(TextContentConfig config)
        {
            _config = config ?? new TextContentConfig();
        }

        internal TextContent(string text)
        {
            _config = new TextContentConfig {Value = text};
        }
    }
}