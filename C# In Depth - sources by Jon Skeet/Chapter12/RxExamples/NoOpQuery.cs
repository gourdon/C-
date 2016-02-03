using System;
using System.ComponentModel;
using System.Linq;

namespace ReactiveExtensions
{
    [Description("Listing 12.13")]
    class NoOpQuery
    {
        static void Main()
        {
            var observable = Observable.Range(0, 10);
            observable.Subscribe(x => Console.WriteLine("Received {0}", x),
                                 e => Console.WriteLine("Error: {0}", e),
                                 () => Console.WriteLine("Finished")); 
        }
    }
}
