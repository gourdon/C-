using System;
using System.Collections;
using System.ComponentModel;
using System.Linq;

namespace Chapter11.Queries
{
    [Description("Listing 11.06")]
    class ExplicitRangeVariable
    {
        static void Main()
        {
            ArrayList list = new ArrayList { "First", "Second", "Third" };
            var strings = from string entry in list
                          select entry.Substring(0, 3);


            foreach (string start in strings)
            {
                Console.WriteLine(start);
            }
        }
    }
}
