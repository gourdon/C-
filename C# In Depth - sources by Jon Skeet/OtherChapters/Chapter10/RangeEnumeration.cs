using System;
using System.ComponentModel;
using System.Linq;

namespace Chapter10
{
    [Description("Listing 10.06")]
    class RangeEnumeration
    {
        static void Main()
        {
            var collection = Enumerable.Range(0, 10);

            foreach (var element in collection)
            {
                Console.WriteLine(element);
            }
        }
    }
}
