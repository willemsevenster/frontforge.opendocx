using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.Converters
{
    internal class PageMarginConverter
    {
        protected PageMarginConverter()
        {
        }

        public static PageMargin PageMarginFromPredefined(PredefinedPageMargins marginType)
        {
            var margins = PageMargins(marginType);

            return new PageMargin
            {
                Top = Int32Value.FromInt32(Convert.ToInt32(margins.top.ToDxa())),
                Left = UInt32Value.FromUInt32(Convert.ToUInt32(margins.left.ToDxa())),
                Right = UInt32Value.FromUInt32(Convert.ToUInt32(margins.right.ToDxa())),
                Bottom = Int32Value.FromInt32(Convert.ToInt32(margins.bottom.ToDxa())),
                Header = UInt32Value.FromUInt32(Convert.ToUInt32(margins.header.ToDxa())),
                Footer = UInt32Value.FromUInt32(Convert.ToUInt32(margins.footer.ToDxa()))
            };
        }

        private static (Unit left, Unit top, Unit right, Unit bottom, Unit header, Unit footer) PageMargins(
            PredefinedPageMargins marginType)
        {
            switch (marginType)
            {
                case PredefinedPageMargins.Normal:
                    return (new Unit(1d, UnitType.inch), new Unit(1d, UnitType.inch), new Unit(1d, UnitType.inch),
                        new Unit(1d, UnitType.inch), new Unit(1.25d, UnitType.cm), new Unit(1.25d, UnitType.cm));

                case PredefinedPageMargins.Narrow:
                    return (new Unit(0.5d, UnitType.inch), new Unit(0.5d, UnitType.inch), new Unit(0.5d, UnitType.inch),
                        new Unit(0.5d, UnitType.inch), new Unit(1.25d, UnitType.cm), new Unit(1.25d, UnitType.cm));

                case PredefinedPageMargins.Moderate:
                    return (new Unit(1.19, UnitType.cm), new Unit(1d, UnitType.inch), new Unit(1.19d, UnitType.cm),
                        new Unit(1d, UnitType.inch), new Unit(1.25d, UnitType.cm), new Unit(1.25d, UnitType.cm));

                default:
                    throw new ArgumentOutOfRangeException(nameof(marginType), marginType, null);
            }
        }
    }
}