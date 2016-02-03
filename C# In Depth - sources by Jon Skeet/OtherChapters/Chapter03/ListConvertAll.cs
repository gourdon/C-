using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter03
{
    [Description("Listing 3.02")]
    class ListConvertAll
    {
        static double TakeSquareRoot(int x)
        {
            return Math.Sqrt(x);
        }

        static void Main()
        {
            List<int> integers = new List<int>();
            integers.Add(1);
            integers.Add(2);
            integers.Add(3);
            integers.Add(4);

            Converter<int, double> converter = TakeSquareRoot;
            List<double> doubles = integers.ConvertAll<double>(converter);

            foreach (double d in doubles)
            {
                Console.WriteLine(d);
            }
        }
    }
}
