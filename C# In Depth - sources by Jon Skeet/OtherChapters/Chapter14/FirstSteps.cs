using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter14
{
    [Description("Listing 14.01")]
    class FirstSteps
    {
        static void Main()
        {
            dynamic items = new List<string> { "First", "Second", "Third" };
            dynamic valueToAdd = "!";
            foreach (dynamic item in items)
            {
               string result = item + valueToAdd;
               Console.WriteLine(result);
            }
        }
    }
}
