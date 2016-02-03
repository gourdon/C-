using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter05
{
    [Description("Listing 5.06")]
    class UglyButCompactCode
    {
        static void Main()
        {
            List<int> x = new List<int>();
            x.Add(5);
            x.Add(10);
            x.Add(15);
            x.Add(20);
            x.Add(25);

            x.ForEach(delegate(int n){Console.WriteLine(Math.Sqrt(n));});
        }
    }
}
