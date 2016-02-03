using System;
using System.ComponentModel;
using System.Linq;

using Chapter11.Model;

namespace Chapter11.Queries
{
    [Description("Listing 11.11")]
    class OrderUsersByLengthOfNameWithLet
    {
        static void Main()
        {
            var query = from user in SampleData.AllUsers
                        let length = user.Name.Length
                        orderby length
                        select new { Name = user.Name, Length = length };

            foreach (var entry in query)
            {
                Console.WriteLine("{0}: {1}", entry.Length, entry.Name);
            }
        }
    }
}
