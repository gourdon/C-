using System;
using System.ComponentModel;
using System.Linq;
using Chapter11.Model;

namespace Chapter11.Queries
{
    [Description("Listing 11.10")]
    class OrderUsersByLengthOfNameWithoutLet
    {
        static void Main()
        {
            var query = from user in SampleData.AllUsers
                        orderby user.Name.Length
                        select user.Name;

            foreach (var name in query)
            {
                Console.WriteLine("{0}: {1}", name.Length, name);
            }
        }
    }
}
