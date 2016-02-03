using System;
using System.Collections.Generic;
using System.Collections;

namespace Chapter11.Queries
{
    public class DateTimeRange : IEnumerable<DateTime>
    {
        private readonly DateTime start;
        private readonly DateTime end;

        public DateTimeRange(DateTime start, DateTime end)
        {
            this.start = start;
            this.end = end;
        }

        public IEnumerator<DateTime> GetEnumerator()
        {
            for (DateTime current = start; current <= end; current = current.AddDays(1))
            {
                yield return current;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
