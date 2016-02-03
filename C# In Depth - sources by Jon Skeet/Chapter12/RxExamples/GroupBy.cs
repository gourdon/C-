using System;
using System.ComponentModel;
using System.Linq;

namespace ReactiveExtensions
{
    [Description("Listing 12.15")]
    class GroupBy
    {
        static void Main()
        {
            var numbers = Observable.Range(0, 10);
            var query = from number in numbers
                        group number by number % 3;
            query.Subscribe(group => group.Subscribe
                (x => Console.WriteLine("Value: {0}; Group: {1}", x, group.Key)));
        }
    }
}
