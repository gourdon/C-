using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace AppendixB
{
    [Description("Listing B.2")]
    class SortedSetView
    {
        static void Main()
        {
            var baseSet = new SortedSet<int> { 1, 5, 12, 20, 25 };
            var view = baseSet.GetViewBetween(10, 20);
            view.Add(14);
            Console.WriteLine(baseSet.Count);
            foreach (int value in view)
            {
                Console.WriteLine(value);
            }
        }
    }
}
