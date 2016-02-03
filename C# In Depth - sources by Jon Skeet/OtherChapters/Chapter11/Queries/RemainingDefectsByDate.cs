using System;
using System.Linq;

using Chapter11.Model;

namespace Chapter11.Queries
{
    class RemainingDefectsByDate
    {
        static void Main()
        {
            var query = from date in new DateTimeRange(SampleData.Start, SampleData.End)
                        select new
                        {
                            Date = date,
                            Count =
                                (from defect in SampleData.AllDefects
                                 // Created before the date we're interested in and either not closed
                                 // or last modified after the date we're interested in
                                 where defect.Created.Date <= date && (defect.Status != Status.Closed || defect.LastModified > date)
                                 select defect).Count()
                        };

            foreach (var grouped in query)
            {
                Console.WriteLine("{0:d}: {1}", grouped.Date, grouped.Count);
            }
        }
    }
}
