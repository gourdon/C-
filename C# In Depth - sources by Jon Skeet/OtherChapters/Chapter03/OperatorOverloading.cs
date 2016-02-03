using System;
using System.ComponentModel;

namespace Chapter03
{
    [Description("Listing 3.05")]
    class OperatorOverloading
    {
        static bool AreReferencesEqual<T>(T first, T second)
            where T : class
        {
            return first == second;
        }

        static void Main()
        {
            string name = "Jon";
            string intro1 = "My name is " + name;
            string intro2 = "My name is " + name;
            Console.WriteLine(intro1 == intro2);
            Console.WriteLine(AreReferencesEqual(intro1, intro2));
        }
    }
}
