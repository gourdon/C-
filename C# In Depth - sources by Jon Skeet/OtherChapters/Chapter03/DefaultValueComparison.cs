using System;
using System.ComponentModel;

namespace Chapter03
{
    [Description("Listing 3.04")]
    class DefaultValueComparison
    {
        static int CompareToDefault<T>(T value)
            where T : IComparable<T>
        {
            return value.CompareTo(default(T));
        }
    
        static void Main()
        {
            Console.WriteLine(CompareToDefault("x"));
            Console.WriteLine(CompareToDefault(10));
            Console.WriteLine(CompareToDefault(0));
            Console.WriteLine(CompareToDefault(-10));
            Console.WriteLine(CompareToDefault(DateTime.MinValue));
        }
    }
}
