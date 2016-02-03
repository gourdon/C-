using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Model;

namespace SqlExamples
{
    class AllDefectsToSubscribersGrouped
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                var query = from defect in context.Defects
                            join subscription in context.NotificationSubscriptions
                                 on defect.Project equals subscription.Project
                                 into groupedSubscriptions
                            select new { Defect = defect, Subscriptions = groupedSubscriptions };

                foreach (var entry in query)
                {
                    Console.WriteLine(entry.Defect.Summary);
                    Console.WriteLine(entry.Subscriptions.Count());
                    foreach (var subscription in entry.Subscriptions)
                    {
                        Console.WriteLine("  {0}", subscription.EmailAddress);
                    }
                }

            }
        }
    }
}
