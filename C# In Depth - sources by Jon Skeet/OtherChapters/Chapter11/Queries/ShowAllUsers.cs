using System;
using System.ComponentModel;
using System.Linq;
using Chapter11.Model;

namespace Chapter11.Queries
{
    [Description("Listing 11.01")]
    class ShowAllUsers
    {
        static void Main()
        {
            var query = from user in SampleData.AllUsers 
                        select user;

            foreach (var user in query)
            {
                Console.WriteLine(user);
            }
        }
    }
}
