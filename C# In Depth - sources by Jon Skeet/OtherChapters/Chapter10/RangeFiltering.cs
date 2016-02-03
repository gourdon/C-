using System;
using System.ComponentModel;
using System.Linq;

namespace Chapter10
{
    [Description("Listing 10.08")]
    class RangeFiltering
    {
        static void Main()
        {
            var collection = Enumerable.Range(0, 10)
                           .Where(x => x % 2 != 0)
                           .Reverse();

            foreach (var element in collection)
            {
                Console.WriteLine(element);
            }
        }
    }
}
