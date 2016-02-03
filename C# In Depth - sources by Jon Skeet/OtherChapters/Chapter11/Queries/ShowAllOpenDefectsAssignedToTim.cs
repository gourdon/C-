using System;
using System.ComponentModel;
using System.Linq;

using Chapter11.Model;

namespace Chapter11.Queries
{
    [Description("Listing 11.07")]
    class ShowAllOpenDefectsAssignedToTim
    {
        static void Main()
        {
            User tim = SampleData.Users.TesterTim;

            var query = from bug in SampleData.AllDefects
                        where bug.Status != Status.Closed
                        where bug.AssignedTo == tim
                        select bug.Summary;

            foreach (var summary in query)
            {
                Console.WriteLine(summary);
            }
        }
    }
}
