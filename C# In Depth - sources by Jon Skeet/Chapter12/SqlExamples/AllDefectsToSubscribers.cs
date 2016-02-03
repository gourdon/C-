using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Model;

namespace SqlExamples
{
    class AllDefectsToSubscribers
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                var query = from defect in context.Defects
                            join subscription in context.NotificationSubscriptions
                                on defect.Project equals subscription.Project
                            select new { defect.Summary, subscription.EmailAddress };

                foreach (var entry in query)
                {
                    Console.WriteLine("{0}: {1}", entry.EmailAddress, entry.Summary);
                }
            }
        }
    }
}
