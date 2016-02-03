using System;
using System.ComponentModel;
using System.Linq;
using Chapter11.Model;

namespace Chapter11.Queries
{
    [Description("Listing 11.17")]
    class GroupDefectsByAssignee
    {
        static void Main()
        {
            var query = from defect in SampleData.AllDefects
                        where defect.AssignedTo != null
                        group defect by defect.AssignedTo;

            foreach (var entry in query)
            {
                Console.WriteLine(entry.Key.Name);
                foreach (var defect in entry)
                {
                    Console.WriteLine("  ({0}) {1}",
                                      defect.Severity,
                                      defect.Summary);
                }
                Console.WriteLine();
            }
        }
    }
}
