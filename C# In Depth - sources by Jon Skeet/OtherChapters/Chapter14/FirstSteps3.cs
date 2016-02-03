using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter14
{
    [Description("Listing 14.03")]
    class FirstSteps3
    {
        static void Main()
        {
            dynamic items = new List<int> { 1, 2, 3 };
            dynamic valueToAdd = 2;
            foreach (dynamic item in items)
            {
               string result = item + valueToAdd;
               Console.WriteLine(result);
            }
        }
    }
}
