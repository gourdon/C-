using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter06
{
    [Description("Listing 6.9")]
    class FakeLinq
    {
        public static IEnumerable<T> Where<T>(IEnumerable<T> source, Predicate<T> predicate)
        {
            if (source == null || predicate == null)
            {
                throw new ArgumentNullException();
            }
            return WhereImpl(source, predicate);
        }

        private static IEnumerable<T> WhereImpl<T>(IEnumerable<T> source, Predicate<T> predicate)
        {
            foreach (T item in source)
            {
                if (predicate(item))
                {
                    yield return item;
                }
            }
        }

        static void Main()
        {
            IEnumerable<string> lines = LineReader.ReadLines("../../FakeLinq.cs");
            Predicate<string> predicate = delegate(string line)
                { return line.StartsWith("using"); };
            foreach (string line in Where(lines, predicate))
            {
                Console.WriteLine(line);
            }
        }
    }
}
