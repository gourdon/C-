using System;
using System.ComponentModel;

namespace Chapter09
{
    [Description("Listing 9.02")]
    class FirstLambdaExpression
    {
        static void Main()
        {
            Func<string, int> returnLength;
            returnLength = (string text) => { return text.Length; };

            Console.WriteLine(returnLength("Hello"));
        }
    }
}
