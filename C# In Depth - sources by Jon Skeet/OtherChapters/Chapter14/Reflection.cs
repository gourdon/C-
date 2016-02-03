using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter14
{
    [Description("Listing 14.11")]
    class Reflection
    {
        private static bool AddConditionallyImpl<T>(IList<T> list, T item)
        {
            if (list.Count > 10)
            {
                list.Add(item);
                return true;
            }
            return false;
        }

        public static bool AddConditionally(dynamic list, dynamic item)
        {
            return AddConditionallyImpl(list, item);
        }
        static void Main()
        {
            object list = new List<string> { "x", "y" };
            object item = "z";
            Console.WriteLine(AddConditionally(list, item));
        }
    }
}
