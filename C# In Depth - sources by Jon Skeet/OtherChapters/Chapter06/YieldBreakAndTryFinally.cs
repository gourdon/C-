using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;

namespace Chapter06
{
    [Description("Listing 6.7")]
    class YieldBreakAndTryFinally
    {
        static IEnumerable<int> CountWithTimeLimit(DateTime limit)
        {
            try
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
            finally
            {
                Console.WriteLine("Stopping!");
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
