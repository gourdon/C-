using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter03
{
    [Description("Listing 3.11")]
    class TypeofOperator
    {
        static internal void DemonstrateTypeof<X>()
        {
            Console.WriteLine(typeof(X));

            Console.WriteLine(typeof(List<>));
            Console.WriteLine(typeof(Dictionary<,>));

            Console.WriteLine(typeof(List<X>));
            Console.WriteLine(typeof(Dictionary<string, X>));

            Console.WriteLine(typeof(List<long>));
            Console.WriteLine(typeof(Dictionary<long, Guid>));
        }

        static void Main()
        {
            DemonstrateTypeof<int>();
        }
    }
}
