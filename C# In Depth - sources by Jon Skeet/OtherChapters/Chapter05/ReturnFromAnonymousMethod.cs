using System;
using System.ComponentModel;

namespace Chapter05
{
    [Description("Listing 5.07")]
    class ReturnFromAnonymousMethod
    {
        static void Main()
        {
            Predicate<int> isEven = delegate(int x)
                { return x % 2 == 0; };

            Console.WriteLine(isEven(1));
            Console.WriteLine(isEven(4));
        }
    }
}
