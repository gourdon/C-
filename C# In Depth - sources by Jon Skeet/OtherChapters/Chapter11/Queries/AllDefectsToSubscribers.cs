using System;
using System.ComponentModel;
using System.Linq;

using Chapter11.Model;

namespace Chapter11.Queries
{
    [Description("Listing 11.12")]
    class AllDefectsToSubscribers
    {
        static void Main()
        {
            var query = from defect in SampleData.AllDefects 
            join subscription in SampleData.AllSubscriptions
                 on defect.Project equals subscription.Project
            select new { defect.Summary, subscription.EmailAddress };

            foreach (var entry in query)
            {
                Console.WriteLine("{0}: {1}", entry.EmailAddress, entry.Summary);
            }
        }
    }
}
