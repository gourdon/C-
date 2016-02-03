using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace AppendixB
{
    [Description("Listing B.1")]
    class CaseInsensitiveDictionary
    {
        static void Main()
        {
            var comparer = StringComparer.OrdinalIgnoreCase;
            var dict = new Dictionary<String, int>(comparer);
            dict["TEST"] = 10;
            Console.WriteLine(dict["test"]);
        }
    }
}
