using System;
using System.ComponentModel;

namespace Chapter09
{
    [Description("Listing 9.12")]
    class ReturnTypeInference
    {
        delegate T MyFunc<T>();

        static void WriteResult<T>(MyFunc<T> function)
        {
            Console.WriteLine(function());
        }

        static void Main()
        {
            WriteResult(delegate { return 5; });
        }
    }
}
