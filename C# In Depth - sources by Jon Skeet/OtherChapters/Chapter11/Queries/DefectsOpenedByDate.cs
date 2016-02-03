using System;
using System.ComponentModel;
using System.Linq;

using Chapter11.Model;

namespace Chapter11.Queries
{
    [Description("Listing 11.14")]
    class DefectsOpenedByDate
    {
        static void Main()
        {
            var dates = new DateTimeRange(SampleData.Start, SampleData.End);

            var query = from date in dates
                        join defect in SampleData.AllDefects 
                             on date equals defect.Created.Date 
                             into joined
                        select new { Date=date, Count=joined.Count() };

            foreach (var grouped in query)
            {
                Console.WriteLine("{0:d}: {1}", grouped.Date, grouped.Count);
            }
        }
    }
}
