using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter14
{
    static class DynamicSumExtensions
    {
        public static T DynamicSum<T>(this IEnumerable<T> source)
        {
            dynamic total = default(T);
            foreach (T element in source)
            {
                total = (T) (total + element);
            }
            return total;
        }
    }

    [Description("Listing 14.12")]
    class DynamicSum
    {
        static void Main()
        {
            byte[] bytes = new byte[] { 1, 2, 3 };
            Console.WriteLine(bytes.DynamicSum());
        }
    }
}
