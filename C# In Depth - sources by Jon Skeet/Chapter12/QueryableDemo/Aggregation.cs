using System.ComponentModel;
using System.Linq;

namespace QueryableDemo
{
    [Description("Listing 12.6")]
    class Aggregation
    {
        static void Main()
        {
            var query = from x in new FakeQuery<string>()
                        where x.StartsWith("abc")
                        select x.Length;

            double mean = query.Average();
        }
    }
}
