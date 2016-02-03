using System;

namespace Chapter09.HigherOrderFunctions
{
    // See http://paulgraham.com/icad.html
    class Accumulator
    {
        static void Main()
        {
            // We want to create a delegate which, which given a number n,
            // returns a function that takes another number i, and returns n
            // incremented by i.
            Func<int, Func<int, int>> accumulatorGenerator = (n => (i => n += i));

            var accumulator = accumulatorGenerator(1);

            Console.WriteLine(accumulator(2)); // Prints 3 (n is now 3)
            Console.WriteLine(accumulator(1)); // Prints 4
        }
    }
}
