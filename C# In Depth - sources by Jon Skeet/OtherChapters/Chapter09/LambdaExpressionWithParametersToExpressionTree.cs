using System;
using System.ComponentModel;
using System.Linq.Expressions;

namespace Chapter09
{
    [Description("Listing 9.09")]
    class LambdaExpressionWithParametersToExpressionTree
    {
        static void Main()
        {
            Expression<Func<string, string, bool>> expression = (x, y) => x.StartsWith(y);

            var compiled = expression.Compile();

            Console.WriteLine(compiled("First", "Second"));
            Console.WriteLine(compiled("First", "Fir"));
        }
    }
}
