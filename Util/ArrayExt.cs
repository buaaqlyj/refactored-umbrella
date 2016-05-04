using System;
using System.Collections.Generic;
using System.Text;

namespace Util
{
    public static class ArrayExt
    {
        public static T[] CombineArray<T>(T[] a1, T[] a2)
        {
            T[] array = new T[a1.Length + a2.Length];
            a1.CopyTo(array, 0);
            a2.CopyTo(array, a1.Length);
            return array;
        }
    }
}
