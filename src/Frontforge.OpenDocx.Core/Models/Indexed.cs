namespace Frontforge.OpenDocx.Core.Models
{
    public class Indexed<T>
    {
        public static implicit operator T(Indexed<T> value)
        {
            return value.Value;
        }

        internal Indexed(int index, int length, T value)
        {
            Index = index;
            Total = length;
            IsFirst = index == 0;
            IsLast = index == length - 1;
            Value = value;
        }

        public int Index { get; }
        public int Total { get; }
        public bool IsFirst { get; }
        public bool IsLast { get; }
        public T Value { get; }
    }
}