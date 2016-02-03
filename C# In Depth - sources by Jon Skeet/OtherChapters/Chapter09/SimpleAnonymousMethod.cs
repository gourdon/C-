using System;
using System.ComponentModel;

namespace Chapter09
{
    [Description("Listing 9.01")]
    class SimpleAnonymousMethod
    {
        static void Main()
        {
            Func<string, int> returnLength;
            returnLength = delegate(string text) { return text.Length; };

            Console.WriteLine(returnLength("Hello"));
        }
    }
}
