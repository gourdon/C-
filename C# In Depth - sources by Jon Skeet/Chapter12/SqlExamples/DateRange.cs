using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SqlExamples
{
    public class DateTimeRange : Range<DateTime>
    {
        readonly TimeSpan step;

        public DateTimeRange(DateTime start, DateTime end)
            : this(start, end, TimeSpan.FromDays(1))
        {
        }

        public DateTimeRange(DateTime start,
                             DateTime end,
                             TimeSpan step)
            : base(start, end)
        {
            this.step = step;
        }

        protected override DateTime GetNextValue(DateTime current)
        {
            return current + step;
        }
    }
}
