using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter03
{
    [Description("Listing 3.12")]
    class GenericTypeReflection
    {
        static void Main()
        {
            string listTypeName = "System.Collections.Generic.List`1";

            Type defByName = Type.GetType(listTypeName);

            Type closedByName = Type.GetType(listTypeName + "[System.String]");
            Type closedByMethod = defByName.MakeGenericType(typeof(string));
            Type closedByTypeof = typeof(List<string>);

            Console.WriteLine(closedByMethod == closedByName);
            Console.WriteLine(closedByName == closedByTypeof);

            Type defByTypeof = typeof(List<>);
            Type defByMethod = closedByName.GetGenericTypeDefinition();

            Console.WriteLine(defByMethod == defByName);
            Console.WriteLine(defByName == defByTypeof);
        }
    }
}
