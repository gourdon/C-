using System.ComponentModel;
using System.Linq;

namespace QueryableDemo
{
    [Description("Listing 12.5")]
    class SimpleQuery
    {
        static void Main()
        {
            var query = from x in new FakeQuery<string>()
                        where x.StartsWith("abc")
                        select x.Length;

            foreach (int i in query) { }
        }
    }
}
