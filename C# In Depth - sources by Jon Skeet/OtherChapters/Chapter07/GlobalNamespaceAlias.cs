using System;
using System.ComponentModel;

class Configuration {}

namespace Chapter07
{
    class Configuration {}

    [Description("Listing 7.07")]
    class GlobalNamespaceAlias
    {
        static void Main()
        {
            Console.WriteLine(typeof(Configuration));
            Console.WriteLine(typeof(global::Configuration));
            Console.WriteLine(typeof(global::Chapter07.GlobalNamespaceAlias));
        }
    }
}