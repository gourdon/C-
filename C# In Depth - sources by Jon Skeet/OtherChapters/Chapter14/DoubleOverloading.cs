using System;
using System.ComponentModel;

namespace Chapter14
{
    [Description("Listing 14.18")]
    class DoubleOverloading
    {
        static void Execute(dynamic x, string y)
        {
            Console.WriteLine("dynamic, string");
        }

        static void Execute(dynamic x, object y)
        {
            Console.WriteLine("dynamic, object");
        }

        static void Main()
        {
            object text = "text";
            dynamic d = 10;
            Execute(d, text);
        }
    }
}
