using System;
using System.ComponentModel;
using System.Dynamic;

namespace Chapter14
{
    [Description("Listing 14.26")]
    class ExpandoDelegate
    {
        static void Main()
        {
            dynamic expando = new ExpandoObject();
            expando.AddOne = (Func<int, int>)(x => x + 1);
            Console.Write(expando.AddOne(10));
        }
    }
}
