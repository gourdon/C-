using System;
using System.ComponentModel;
using System.Linq;

namespace Chapter10
{
    [Description("Listing 10.07")]
    class RangeReversal
    {
        static void Main()
        {
            var collection = Enumerable.Range(0, 10)
                           .Reverse();

            foreach (var element in collection)
            {
                Console.WriteLine(element);
            }
        }
    }
}
