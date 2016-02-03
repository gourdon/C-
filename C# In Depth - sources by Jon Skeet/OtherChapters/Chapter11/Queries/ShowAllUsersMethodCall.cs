using System;
using System.ComponentModel;
using System.Linq;
using Chapter11.Model;

namespace Chapter11.Queries
{
    [Description("Listing 11.02")]
    class ShowAllUsersMethodCall
    {
        static void Main()
        {
            var query = SampleData.AllUsers.Select(user => user);

            foreach (var user in query)
            {
                Console.WriteLine(user);
            }
        }
    }
}
