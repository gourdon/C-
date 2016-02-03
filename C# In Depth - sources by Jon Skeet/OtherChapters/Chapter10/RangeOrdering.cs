using System;
using System.ComponentModel;
using System.Linq;

namespace Chapter10
{
    [Description("Listing 10.10")]
    class RangeOrdering
    {
        static void Main()
        {
            var collection = Enumerable.Range(-5, 11)
                                       .Select(x => new { Original = x, Square = x * x })
                                       .OrderBy(x => x.Square)
                                       .ThenBy(x => x.Original);

            foreach (var element in collection)
            {
                Console.WriteLine(element);
            }
        }
    }
}
