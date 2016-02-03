using System;
using System.ComponentModel;
using System.Linq;

using Chapter11.Model;

namespace Chapter11.Queries
{
    [Description("Listing 11.19")]
    class GroupDefectsByAssigneeWithContinuation
    {
        static void Main()
        {
            var query = from defect in SampleData.AllDefects
                        where defect.AssignedTo != null
                        group defect by defect.AssignedTo into grouped
                        select new { Assignee = grouped.Key, Count = grouped.Count() };

            foreach (var entry in query)
            {
                Console.WriteLine("{0}: {1}",
                                  entry.Assignee.Name,
                                  entry.Count);
            }
        }
    }
}
