using System;
using System.ComponentModel;
using System.Linq.Expressions;

namespace Chapter09
{
    [Description("Listing 9.07")]
    class CompiledExpressionTree
    {
        static void Main()
        {
            Expression firstArg = Expression.Constant(2);
            Expression secondArg = Expression.Constant(3);
            Expression add = Expression.Add(firstArg, secondArg);

            Func<int> compiled = Expression.Lambda<Func<int>>(add).Compile();
            Console.WriteLine(compiled());
        }
    }
}
