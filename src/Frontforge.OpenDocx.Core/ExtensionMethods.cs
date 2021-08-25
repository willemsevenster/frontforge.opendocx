using System.Collections.Generic;
using System.Linq;
using System.Text;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core
{
    public static class ExtensionMethods
    {
        public static IEnumerable<Indexed<T>> AsIndexed<T>(this IEnumerable<T> values)
        {
            var enumerable = values.ToList();
            var count = enumerable.Count;
            return enumerable.Select((x, i) => new Indexed<T>(i, count, x));
        }


        public static string StringTogether(this IEnumerable<Indexed<string>> values, string separator,
            string lastSeparator = null)
        {
            var result = new StringBuilder();

            foreach (var value in values)
            {
                if (value.IsFirst)
                {
                    result.Append(value.Value);
                    continue;
                }

                if (!value.IsLast)
                {
                    result.Append(separator ?? string.Empty);
                    result.Append(value.Value);
                    continue;
                }

                result.Append(lastSeparator ?? separator ?? string.Empty);
                result.Append(value.Value);
            }

            return result.ToString();
        }
    }
}