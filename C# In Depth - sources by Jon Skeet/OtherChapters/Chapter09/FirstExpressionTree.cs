using System;
using System.ComponentModel;
using System.Linq.Expressions;

namespace Chapter09
{
    [Description("Listing 9.06")]
    class FirstExpressionTree
    {
        static void Main()
        {
            Expression firstArg = Expression.Constant(2);
            Expression secondArg = Expression.Constant(3);
            Expression add = Expression.Add(firstArg, secondArg);

            Console.WriteLine(add);
        }
    }
}
