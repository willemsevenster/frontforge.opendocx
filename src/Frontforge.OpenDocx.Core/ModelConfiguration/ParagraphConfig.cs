using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.Converters;
using Frontforge.OpenDocx.Core.Models;
using CheckBox = DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox;
using Checked = DocumentFormat.OpenXml.Office2010.Word.Checked;

namespace Frontforge.OpenDocx.Core.ModelConfiguration
{
    internal class ParagraphConfig
    {
        #region properties

        public bool? Bold { get; set; }

        public Unit FontSize { get; set; }

        public Unit SpacingBefore { get; set; }

        public Unit SpacingAfter { get; set; }

        public HorizontalAlignment? HorizontalAlignment { get; set; }

        public List<IContent> Contents { get; } = new List<IContent>();

        public ParagraphBorders Borders { get; } = new ParagraphBorders();
        public bool? Bullets { get; set; }

        #endregion

        #region members

        public ParagraphProperties GetParagraphProperties()
        {
            var result = new ParagraphProperties();

            if (HorizontalAlignment.HasValue)
            {
                result.Justification = AlignmentConverter.CreateJustification(HorizontalAlignment.Value);
            }

            result.SpacingBetweenLines = new SpacingBetweenLines();

            if (SpacingBefore != null)
            {
                result.SpacingBetweenLines.Before = SpacingBefore;
            }

            if (SpacingAfter != null)
            {
                result.SpacingBetweenLines.After = SpacingAfter;
            }

            result.ParagraphBorders = Borders;

            if (Bullets == true)
            {
                result.ParagraphStyleId = new ParagraphStyleId {Val = "ListParagraph"};
                result.NumberingProperties = new NumberingProperties
                {
                    NumberingLevelReference = new NumberingLevelReference {Val = 0},
                    NumberingId = new NumberingId {Val = 1}
                };
                //var numberingProperties = new NumberingProperties(new NumberingLevelReference() { Val = 1 });
                //var spacingBetweenLines1 = new SpacingBetweenLines() { After = "0" };  // Get rid of space between bullets
                //var indentation = new Indentation() { Left = "360", Hanging = "180" };  // correct indentation 

                //var paragraphMarkRunProperties1 = new ParagraphMarkRunProperties(
                //    new RunFonts() {Ascii = "Symbol", HighAnsi = "Symbol"}
                //);

                //result.Append(numberingProperties, spacingBetweenLines1, indentation, paragraphMarkRunProperties1);
            }

            return result;
        }

        public RunProperties GetRunProperties()
        {
            var result = new RunProperties();

            if (Bold.HasValue)
            {
                result.Bold = new Bold {Val = Bold};
            }

            if (FontSize != null)
            {
                var mrp = new[]
                {
                    new ParagraphMarkRunProperties(
                        new FontSize {Val = FontSize}
                    )
                };

                result.Append(mrp.AsEnumerable());
            }

            return result;
        }

        public IEnumerable<Run> GetRuns()
        {
            var runProperties = GetRunProperties();

            foreach (var content in Contents)
            {
                if (content is TextContent text)
                {
                    yield return text.GetRun(runProperties);
                }

                if (content is CheckboxControl checkbox)
                {
                    var run = new Run {RunProperties = runProperties.CloneNode()};

                    var cb = new CheckBox(
                        new Checked {Val = checkbox.IsChecked ? OnOffValues.One : OnOffValues.Zero},
                        new CheckedState {Val = "0052", Font = "Wingdings 2"},
                        new UncheckedState {Val = "00A3", Font = "Wingdings 2"}
                    );

                    var cbSdt = new SdtBlock(
                        new SdtProperties(
                            new Lock {Val = LockingValues.ContentLocked},
                            new Appearance {Val = SdtAppearance.Hidden}, cb),
                        new SdtContentBlock(
                            new Run(new SymbolChar
                            {
                                Font = "Wingdings 2",
                                Char = checkbox.IsChecked ? "F052" : "F0A3"
                            }))
                    );

                    run.AppendChild(cbSdt);

                    yield return run;

                    if (!string.IsNullOrWhiteSpace(checkbox.Label))
                    {
                        yield return new Run(
                            new Text(checkbox.Label)
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            })
                        {
                            RunProperties = runProperties.CloneNode()
                        };
                    }
                }
            }
        }

        #endregion
    }
}