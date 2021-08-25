using System;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.Converters
{
    internal static class AlignmentConverter
    {
        public static Justification CreateJustification(HorizontalAlignment alignment)
        {
            return new Justification {Val = Convert(alignment)};
        }

        public static JustificationValues Convert(HorizontalAlignment alignment)
        {
            return alignment switch
            {
                HorizontalAlignment.Left => JustificationValues.Left,
                HorizontalAlignment.Center => JustificationValues.Center,
                HorizontalAlignment.Right => JustificationValues.Right,
                HorizontalAlignment.Justified => JustificationValues.Both,
                _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null)
            };
        }
    }
}