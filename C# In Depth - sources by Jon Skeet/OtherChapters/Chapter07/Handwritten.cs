using System;
using System.ComponentModel;

namespace Chapter07
{
    [Description("Listing 7.02")]
    partial class PartialMethodDemo
    {
        partial void OnConstructorEnd()
        {
            Console.WriteLine("Manual code");
        }
    }
}
