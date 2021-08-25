using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Frontforge.OpenDocx.Core
{
    internal static class Extensions
    {
        public static T CloneNode<T>(this T element, bool deep = true)
            where T : OpenXmlElement
        {
            return (T) element.CloneNode(deep);
        }

        public static T Chain<T>(this T obj, Action<T> action)
        {
            action(obj);
            return obj;
        }

        public static T Chain<T, T1>(this T obj, T1 arg, Action<T, T1> action)
        {
            action(obj, arg);
            return obj;
        }

        public static T Chain<T, T1, T2>(this T obj, T1 arg1, T2 arg2, Action<T, T1, T2> action)
        {
            action(obj, arg1, arg2);
            return obj;
        }

        public static BorderType SetBorderColor(this BorderType obj, string color)
        {
            obj.Color = color;
            return obj;
        }
    }
}