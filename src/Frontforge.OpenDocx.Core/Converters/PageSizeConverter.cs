using System;
using Frontforge.OpenDocx.Core.Models;
using PageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;

namespace Frontforge.OpenDocx.Core.Converters
{
    internal static class PageSizeConverter
    {
        public static PageSize PageSize(Models.PageSize pageSize = Models.PageSize.A4,
            PageOrientation orientation = PageOrientation.Portrait)
        {
            var sizes = _pageSize(pageSize, orientation);

            return new PageSize
            {
                Width = Convert.ToUInt32(sizes.width.ToDxa()),
                Height = Convert.ToUInt32(sizes.height.ToDxa()),
                Code = (ushort) pageSize
            };
        }

        private static (Unit width, Unit height) _pageSize(Models.PageSize pageSize, PageOrientation orientation)
        {
            (Unit width, Unit height) result;

            switch (pageSize)
            {
                case Models.PageSize.Letter:
                    result = (new Unit(21.59d, UnitType.cm), new Unit(27.94d, UnitType.cm));
                    break;

                case Models.PageSize.A5:
                    result = (new Unit(14.8d, UnitType.cm), new Unit(21d, UnitType.cm));
                    break;

                case Models.PageSize.A4:
                    result = (new Unit(21d, UnitType.cm), new Unit(29.7d, UnitType.cm));
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(pageSize), pageSize, null);
            }

            if (orientation == PageOrientation.Portrait)
            {
                return result;
            }

            var tmp = result.height;
            result.height = result.width;
            result.width = tmp;

            return result;
        }
    }
}