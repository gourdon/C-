using System;
using System.ComponentModel;
using System.Linq;

namespace ReactiveExtensions
{
    [Description("Listing 12.14")]
    class SelectAndWhere
    {
        static void Main()
        {
            var numbers = Observable.Range(0, 10);
            var query = from number in numbers
                        where number % 2 == 0
                        select number * number;
            query.Subscribe(Console.WriteLine);
        }
    }
}
