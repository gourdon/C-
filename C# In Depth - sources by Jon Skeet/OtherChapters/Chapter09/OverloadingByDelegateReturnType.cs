using System;
using System.ComponentModel;

namespace Chapter09
{
    [Description("Listing 9.16")]
    class OverloadingByDelegateReturnType
    {
        static void Execute(Func<int> action)
        {
            Console.WriteLine("action returns an int: " + action());
        }

        static void Execute(Func<double> action)
        {
            Console.WriteLine("action returns a double: " + action());
        }

        static void Main()
        {
            Execute(() => 1);
        }
    }
}
