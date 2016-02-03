using System;
using System.ComponentModel;

namespace Chapter02
{
    [Description("Listing 2.5")]
    class LambdaExpression
    {
        static void Main()
        {
            Func<int, int, string> func = (x, y) => (x * y).ToString();
            Console.WriteLine(func(5, 20));
        }
    }
}
