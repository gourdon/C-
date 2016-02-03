using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter14
{
    [Description("Listing 14.14")]
    class DuckTypedCount
    {
        static void PrintCount(IEnumerable collection)
        {
            dynamic d = collection;
            int count = d.Count;
            Console.WriteLine(count);
        }

        static void Main()
        {
            PrintCount(new BitArray(10));
            PrintCount(new HashSet<int> { 3, 5 });
            PrintCount(new List<int> { 1, 2, 3 });
        }
    }
}
