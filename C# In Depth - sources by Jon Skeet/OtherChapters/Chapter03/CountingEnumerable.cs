using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter03
{
    [Description("Listing 3.10")]
    class CountingEnumerableExample
    {
        class CountingEnumerable : IEnumerable<int>
        {
            public IEnumerator<int> GetEnumerator()
            {
                return new CountingEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        class CountingEnumerator : IEnumerator<int>
        {
            int current = -1;

            public bool MoveNext()
            {
                current++;
                return current < 10;
            }

            public int Current
            {
                get { return current; }
            }

            object IEnumerator.Current
            {
                get { return Current; }
            }

            public void Reset()
            {
                throw new NotSupportedException();
            }

            public void Dispose()
            {
            }
        }

        static void Main()
        {
            CountingEnumerable counter = new CountingEnumerable();
            foreach (int x in counter)
            {
                Console.WriteLine(x);
            }
        }
    }
}