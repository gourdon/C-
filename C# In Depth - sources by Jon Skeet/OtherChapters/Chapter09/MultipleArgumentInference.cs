using System;
using System.ComponentModel;

namespace Chapter09
{
    [Description("Listing 9.14")]
    class MultipleArgumentInference
    {
        static void PrintType<T>(T first, T second)
        {
            Console.WriteLine(typeof(T));
        }

        static void Main()
        {
            PrintType(1, new object());
        }
    }
}
