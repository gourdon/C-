using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter14
{
    [Description("Listing 14.02")]
    class FirstSteps2
    {
        static void Main()
        {
            dynamic items = new List<string> { "First", "Second", "Third" };
            dynamic valueToAdd = 2;
            foreach (dynamic item in items)
            {
               string result = item + valueToAdd;
               Console.WriteLine(result);
            }
        }
    }
}
