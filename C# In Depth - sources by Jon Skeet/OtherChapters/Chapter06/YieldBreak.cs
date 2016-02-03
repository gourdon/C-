using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;

namespace Chapter06
{
    [Description("Listing 6.6")]
    class YieldBreak
    {
        static IEnumerable<int> CountWithTimeLimit(DateTime limit)
        {
            for (int i = 1; i <= 100; i++)
            {
                if (DateTime.Now >= limit)
                {
                    yield break;
                }
                yield return i;
            }
        }

        static void Main()
        {
            DateTime stop = DateTime.Now.AddSeconds(2);
            foreach (int i in CountWithTimeLimit(stop))
            {
                Console.WriteLine("Received {0}", i);
                Thread.Sleep(300);
            }
        }
    }
}
