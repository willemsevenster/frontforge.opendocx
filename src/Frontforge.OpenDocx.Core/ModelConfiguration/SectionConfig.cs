using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.Converters;
using Frontforge.OpenDocx.Core.Models;
using PageSize = Frontforge.OpenDocx.Core.Models.PageSize;

namespace Frontforge.OpenDocx.Core.ModelConfiguration
{
    internal class SectionConfig
    {
        public PageSize PageSize { get; set; } = PageSize.A4;

        public PredefinedPageMargins PageMargins { get; set; } = PredefinedPageMargins.Normal;

        public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;

        public List<ContentElement> Contents { get; } = new List<ContentElement>();

        public List<ContentElement> Header { get; } = new List<ContentElement>();

        public List<ContentElement> Footer { get; } = new List<ContentElement>();

        internal SectionProperties SectionProperties(MainDocumentPart mainPart)
        {
            var result = new SectionProperties(
                PageSizeConverter.PageSize(PageSize, Orientation),
                PageMarginConverter.PageMarginFromPredefined(PageMargins));

            if (Header?.Any() == true)
            {
                CreateHeader(result, mainPart);
            }

            if (Footer?.Any() == true)
            {
                CreateFooter(result, mainPart);
            }

            return result;
        }


        private void CreateHeader(SectionProperties sectionProperties, MainDocumentPart mainPart)
        {
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            var partId = mainPart.GetIdOfPart(headerPart);
            headerPart.Header = new Header(Header.Where(x => x != null).AsIndexed()
                .Select(x => x.Value.Render(x.Index, x.IsFirst, x.IsLast)));

            sectionProperties.AppendChild(new HeaderReference {Id = partId, Type = HeaderFooterValues.Default});

            headerPart.Header.Save();
        }

        private void CreateFooter(SectionProperties sectionProperties, MainDocumentPart mainPart)
        {
            var footerPart = mainPart.AddNewPart<FooterPart>();
            var partId = mainPart.GetIdOfPart(footerPart);
            footerPart.Footer = new Footer(Footer.Where(x => x != null).AsIndexed()
                .Select(x => x.Value.Render(x.Index, x.IsFirst, x.IsLast)));

            sectionProperties.AppendChild(new FooterReference {Id = partId, Type = HeaderFooterValues.Default});

            footerPart.Footer.Save();
        }
    }
}