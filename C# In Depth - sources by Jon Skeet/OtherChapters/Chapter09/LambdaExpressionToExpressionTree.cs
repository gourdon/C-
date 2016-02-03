using System;
using System.ComponentModel;
using System.Linq.Expressions;

namespace Chapter09
{
    [Description("Listing 9.08")]
    class LambdaExpressionToExpressionTree
    {
        static void Main()
        {
            Expression<Func<int>> return5 = () => 5;
            Func<int> compiled = return5.Compile();
            Console.WriteLine(compiled());
        }
    }
}
