using System;
using System.Linq;

using Model;

namespace SqlExamples
{
    class ShowDefectAndAssignee
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                var query = from defect in context.Defects
                            select new { defect.Summary, Assignee=defect.AssignedTo.Name };

                foreach (var entry in query)
                {
                    Console.WriteLine("{0}: {1}", entry.Assignee, entry.Summary);
                }
            }
        }
    }
}
