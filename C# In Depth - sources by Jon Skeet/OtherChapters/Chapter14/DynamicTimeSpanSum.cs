using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter14
{
    static class TimeSpanExtensions
    {
        internal static TimeSpan Hours(this int hours)
        {
            return TimeSpan.FromHours(hours);
        }

        internal static TimeSpan Minutes(this int minutes)
        {
            return TimeSpan.FromMinutes(minutes);
        }

        internal static TimeSpan Seconds(this int seconds)
        {
            return TimeSpan.FromSeconds(seconds);
        }
    }

    [Description("Listing 14.13")]
    class DynamicTimeSpanSum
    {
        static void Main()
        {
            var times = new List<TimeSpan>
            {                 
                2.Hours(), 25.Minutes(), 30.Seconds(),
                45.Seconds(), 40.Minutes()
            };
            Console.WriteLine(times.DynamicSum());
        }
    }
}
