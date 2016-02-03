using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

using Chapter11.Model;

namespace Chapter11.Queries
{
    [Description("Listing 11.04")]
    class ShowAllUserNames
    {
        static void Main()
        {
            IEnumerable<string> query = from user in SampleData.AllUsers
                                        select user.Name;

            foreach (string name in query)
            {
                Console.WriteLine(name);
            }
        }
    }
}
