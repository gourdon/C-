using System;
using System.ComponentModel;

namespace Chapter09
{
    [Description("Listing 9.03")]
    class ConciseLambdaExpression
    {
        static void Main()
        {
            Func<string, int> returnLength;
            returnLength = text => text.Length;

            Console.WriteLine(returnLength("Hello"));
        }
    }
}
