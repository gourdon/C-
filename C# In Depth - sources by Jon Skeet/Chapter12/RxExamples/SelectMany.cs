using System;
using System.ComponentModel;
using System.Linq;

namespace ReactiveExtensions
{
    [Description("Listing 12.16")]
    class SelectMany
    {
        static void Main()
        {
            var query = from x in Observable.Range(1, 3)
                        from y in Observable.Range(1, x)
                        select new { x, y };
            query.Subscribe(Console.WriteLine);
        }
    }
}
