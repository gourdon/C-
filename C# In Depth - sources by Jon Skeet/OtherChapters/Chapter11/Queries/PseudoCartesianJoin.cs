using System;
using System.Linq;

namespace Chapter11.Queries
{
    class PseudoCartesianJoin
    {
        static void Main()
        {
            var query = from first in Enumerable.Range(1, 5)
                        from second in Enumerable.Range(1, first)
                        select new { first, second };

            foreach (var item in query)
            {
                Console.WriteLine(item);
            }
        }
    }
}
