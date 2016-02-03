using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;

namespace Chapter14
{
    [Description("Listing 14.25")]
    class SimpleExpando
    {
        static void Main()
        {
            dynamic expando = new ExpandoObject();
            IDictionary<string, object> dictionary = expando;
            expando.First = "value set dynamically";
            Console.WriteLine(dictionary["First"]);

            dictionary["Second"] = "value set with dictionary";
            Console.WriteLine(expando.Second);
        }
    }
}
