using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter03.Variance
{
    [Description("Extra code for P104 (conversions for variance)")]
    class EnumeratorConversion
    {
        static void Main()
        {
            IEnumerable<string> strings = new string[] { "First", "Second", "Third" };

            // Not allowed!
            // IEnumerable<object> objects = strings;

            // Convert the IEnumerable<string>
            IEnumerable<object> objects = ConvertEnumerable<string, object>(strings);
            foreach (object o in objects)
            {
                Console.WriteLine(o);
            }

            // Convert an IEnumerator<string>
            using (IEnumerator<string> stringEnumerator = strings.GetEnumerator())
            {
                // Not allowed!
                // IEnumerator<object> objectEnumerator = stringEnumerator;
                IEnumerator<object> objectEnumerator = ConvertEnumerator<string, object>(stringEnumerator);
                while (objectEnumerator.MoveNext())
                {
                    Console.WriteLine(objectEnumerator.Current);
                }
            }
        }

        public static IEnumerable<TDest> ConvertEnumerable<TSource, TDest>(IEnumerable<TSource> source)
            where TSource : TDest
        {
            foreach (TSource element in source)
            {
                // Implicit conversion from TSource to TDest
                yield return element;
            }
        }

        public static IEnumerator<TDest> ConvertEnumerator<TSource, TDest>(IEnumerator<TSource> source)
            where TSource : TDest
        {
            // Note: assume caller will dispose of the enumerator
            while (source.MoveNext())
            {
                // Implicit conversion from TSource to TDest
                yield return source.Current;
            }
        }
    }
}
