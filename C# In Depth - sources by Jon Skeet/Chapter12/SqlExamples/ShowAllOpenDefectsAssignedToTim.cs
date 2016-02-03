using System;
using System.ComponentModel;
using System.Linq;

using Model;

namespace SqlExamples
{
    [Description("Listing 12.1")]
    class ShowAllOpenDefectsAssignedToTim
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                User tim = (from user in context.Users
                            where user.Name=="Tim Trotter"
                            select user).Single();

                var query = from defect in context.Defects
                            where defect.Status != Status.Closed
                            where defect.AssignedTo == tim
                            select defect.Summary;

                foreach (var summary in query)
                {
                    Console.WriteLine(summary);
                }
            }
        }
    }
}
