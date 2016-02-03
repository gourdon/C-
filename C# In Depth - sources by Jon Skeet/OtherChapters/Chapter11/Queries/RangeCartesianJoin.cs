using System;
using System.ComponentModel;
using System.Linq;

namespace Chapter11.Queries
{
    [Description("Listing 11.16")]
    class RangeCartesianJoin
    {
        static void Main()
        {
            var query = from left in Enumerable.Range(1, 4)
                        from right in Enumerable.Range(11, left)
                        select new { Left = left, Right = right };

            foreach (var pair in query)
            {
                Console.WriteLine("Left={0}; Right={1}", pair.Left, pair.Right);
            }
        }
    }
}
