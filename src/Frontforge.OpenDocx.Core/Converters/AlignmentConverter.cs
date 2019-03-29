using System;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.Converters
{
    internal static class AlignmentConverter
    {
        #region members

        public static Justification CreateJustification(HorizontalAlignment alignment)
        {
            return new Justification {Val = Convert(alignment)};
        }

        public static JustificationValues Convert(HorizontalAlignment alignment)
        {
            switch (alignment)
            {
                case HorizontalAlignment.Left:
                    return JustificationValues.Left;
                case HorizontalAlignment.Center:
                    return JustificationValues.Center;
                case HorizontalAlignment.Right:
                    return JustificationValues.Right;
                case HorizontalAlignment.Justified:
                    return JustificationValues.Both;
                default:
                    throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null);
            }
        }

        #endregion
    }
}