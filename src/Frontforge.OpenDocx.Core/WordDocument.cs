using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.Builders;
using Frontforge.OpenDocx.Core.ModelConfiguration;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core
{
    public abstract class WordDocument
    {

        private readonly WordDocumentConfig _config = new WordDocumentConfig();

        private static void AddNumberingStyles(MainDocumentPart mainPart)
        {
            var part = mainPart.NumberingDefinitionsPart;

            if (part == null)
            {
                part = mainPart.AddNewPart<NumberingDefinitionsPart>();
                new Numbering().Save(part);
            }

            var abstractNumberId =
                (part.Numbering.Elements<AbstractNum>().Max(x => x.AbstractNumberId?.Value) ?? 0) + 1;

            var abstractLevel = new Level(
                new StartNumberingValue {Val = 1},
                new NumberingFormat {Val = NumberFormatValues.Bullet},
                new LevelText {Val = ""},
                new LevelJustification {Val = LevelJustificationValues.Left},
                new ParagraphProperties
                {
                    Indentation = new Indentation
                    {
                        Left = new Unit(0.25, UnitType.inch),
                        Hanging = new Unit(0.125, UnitType.inch)
                    }
                },
                new ParagraphMarkRunProperties(
                    new RunFonts
                    {
                        Ascii = "Symbol",
                        HighAnsi = "Symbol",
                        Hint = FontTypeHintValues.Default
                    }
                )
            )
            {
                LevelIndex = 0
            };

            var abstractNum1 = new AbstractNum(abstractLevel)
            {
                AbstractNumberId = abstractNumberId,
                MultiLevelType = new MultiLevelType {Val = MultiLevelValues.HybridMultilevel}
            };

            part.Numbering.AppendChild(abstractNum1);

            part.Numbering.AppendChild(
                new NumberingInstance
                {
                    NumberID = abstractNumberId,
                    AbstractNumId = new AbstractNumId
                    {
                        Val = abstractNumberId
                    }
                }
            );

            part.Numbering.Save();
        }

        private static void AddSettings(MainDocumentPart mainPart)
        {
            var settingsPart = mainPart.DocumentSettingsPart;

            if (settingsPart == null)
            {
                settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();

                settingsPart.Settings = new Settings(
                    new Compatibility(
                        new CompatibilitySetting
                        {
                            Name = new EnumValue<CompatSettingNameValues>
                                (CompatSettingNameValues.CompatibilityMode),
                            Val = new StringValue("15"),
                            Uri = new StringValue("http://schemas.microsoft.com/office/word")
                        }
                    ),
                    new DefaultTabStop {Val = Convert.ToInt16(new Unit(0.25d, UnitType.inch).ToDxa())}
                );
            }

            settingsPart.Settings.Save();
        }

        private static void AddStyles(MainDocumentPart mainPart)
        {
            var part = mainPart.StyleDefinitionsPart;

            if (part == null)
            {
                part = mainPart.AddNewPart<StyleDefinitionsPart>();
                new Styles().Save(part);
            }

            var latent = new LatentStyles(
                new LatentStyleExceptionInfo
                    {Name = "Normal", UiPriority = 0, PrimaryStyle = OnOffValue.FromBoolean(true)},
                new LatentStyleExceptionInfo
                {
                    Name = "No List", SemiHidden = OnOffValue.FromBoolean(true),
                    UnhideWhenUsed = OnOffValue.FromBoolean(true)
                },
                new LatentStyleExceptionInfo
                    {Name = "List Paragraph", UiPriority = 34, PrimaryStyle = OnOffValue.FromBoolean(true)})
            {
                DefaultLockedState = OnOffValue.FromBoolean(false),
                DefaultUiPriority = 99,
                DefaultSemiHidden = OnOffValue.FromBoolean(false),
                DefaultUnhideWhenUsed = OnOffValue.FromBoolean(false),
                DefaultPrimaryStyle = OnOffValue.FromBoolean(false)
            };

            var normalStyle = new Style(
                new PrimaryStyle()
            )
            {
                StyleId = "Normal",
                StyleName = new StyleName {Val = "Normal"}
            };

            var numbering = new Style(
                new StyleName {Val = "No List"},
                new UIPriority {Val = 99},
                new SemiHidden(),
                new UnhideWhenUsed()
            )
            {
                Type = StyleValues.Numbering,
                Default = OnOffValue.FromBoolean(true),
                StyleId = "NoList"
            };


            var listParagraphStyle = new Style(
                new PrimaryStyle(),
                new ParagraphProperties(
                    new Indentation {Left = new Unit(0.25, UnitType.inch)},
                    new ContextualSpacing()
                ))
            {
                Type = StyleValues.Paragraph,
                StyleId = "ListParagraph",
                StyleName = new StyleName {Val = "List Paragraph"},
                BasedOn = new BasedOn {Val = "Normal"},
                UIPriority = new UIPriority {Val = 34}
            };

            part.Styles.Append(latent, numbering, normalStyle, listParagraphStyle);

            part.Styles.Save();
        }

        protected void AddImageMedia(byte[] content, string contentType, string name)
        {
            _config.ImageMedia.Add(new ImageContentItem
            {
                Name = name, Content = content, ContentType = contentType
            });
        }

        private void AddImages(MainDocumentPart mainPart)
        {
            if (!_config.ImageMedia.Any()) return;
            foreach (var imageContent in _config.ImageMedia)
            {
                var imagePart = mainPart.AddNewPart<ImagePart>(imageContent.ContentType, imageContent.Name);
                using var ms = new MemoryStream(imageContent.Content);
                imagePart.FeedData(ms);
                ms.Close();
            }
        }

        protected WordDocument AddSection(Section page)
        {
            return this.Chain(p => p._config.Sections.Add(page));
        }

        protected TextContentBuilder Break()
        {
            return new TextContentBuilder(null);
        }

        protected CellBuilder Cell()
        {
            return new CellBuilder();
        }

        protected CellBuilder Cell(params ContentElement[] contents)
        {
            return new CellBuilder().Add(contents);
        }

        protected CellBuilder Cell(string text, bool bold = false)
        {
            return new CellBuilder().Add(Par(text).Bold(bold).SpacingAfter(0));
        }

        protected CellBuilder Cell(string text, HorizontalAlignment alignment, bool bold = false)
        {
            return new CellBuilder().Add(Par(text, alignment).Bold(bold).SpacingAfter(0));
        }

        protected CheckboxBuilder Checkbox(string text)
        {
            return new CheckboxBuilder().Label(text);
        }

        protected ImageContentBuilder Image(string name)
        {
            return new ImageContentBuilder(name);
        }
        
        protected ParagraphBuilder Par()
        {
            return new ParagraphBuilder();
        }

        protected ParagraphBuilder Par(string text)
        {
            return new ParagraphBuilder().Add(text);
        }

        protected ParagraphBuilder Par(string text, HorizontalAlignment alignment)
        {
            return new ParagraphBuilder().Alignment(alignment).Add(text);
        }

        protected ParagraphBuilder Par(params ContentElement[] contents)
        {
            return new ParagraphBuilder().Add(contents);
        }

        protected RowBuilder Row(params Cell[] cells)
        {
            return new RowBuilder().Add(cells);
        }

        public void Save(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
            var mainPart = document.AddMainDocumentPart();

            AddSettings(mainPart);
            AddNumberingStyles(mainPart);
            AddStyles(mainPart);
            AddImages(mainPart);

            foreach (var section in _config.Sections.Where(x => x != null).AsIndexed())
            {
                var sectionPart = section.Value.Render(section.Index, section.IsFirst, section.IsLast, mainPart);
                new Document(sectionPart).Save(mainPart);
            }

            document.Save();
        }

        protected SectionBuilder Section()
        {
            return new SectionBuilder();
        }

        protected TableBuilder Table(params Row[] rows)
        {
            return new TableBuilder().Add(rows);
        }

        protected TextContentBuilder Text(string text)
        {
            return new TextContentBuilder(text);
        }
    }
}