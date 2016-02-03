using System;
using System.ComponentModel;
using System.Linq;

using Chapter11.Model;

namespace Chapter11.Queries
{
    [Description("Listing 11.18")]
    class GroupDefectsByAssigneeWithProjection
    {
        static void Main()
        {
            var query = from defect in SampleData.AllDefects
                        where defect.AssignedTo != null
                        group defect.Summary by defect.AssignedTo;

            foreach (var entry in query)
            {
                Console.WriteLine(entry.Key.Name);
                foreach (var summary in entry)
                {
                    Console.WriteLine("  {0}", summary);
                }
                Console.WriteLine();
            }
        }
    }
}
