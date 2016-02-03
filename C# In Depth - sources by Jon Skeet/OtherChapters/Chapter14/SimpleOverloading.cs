using System;
using System.ComponentModel;

namespace Chapter14
{
    [Description("Listing 14.16")]
    class SimpleOverloading
    {
        static void Execute(string x)
        {
           Console.WriteLine("String overload");
        }

        static void Execute(dynamic x)
        {
           Console.WriteLine("Dynamic overload");
        }

        static void Main()
        {
            dynamic text = "text";
            Execute(text);
            dynamic number = 10;
            Execute(number);
        }
    }
}
