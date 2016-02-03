using System;

namespace Chapter09.HigherOrderFunctions
{
    static class DoubleApplicationExtensions
    {
        // Takes a function which takes a single parameter and returns the same type.
        // Returns a function which takes a single parameter and returns the same type,
        // by applying the original function to the parameter, then applying it again
        // to the result.
        public static Func<T, T> ApplyTwice<T>(this Func<T, T> original)
        {
            return x => original(original(x));
        }
    }

    class DoubleApplication
    {
        static void Main()
        {
            // Adds 1
            Func<int, int> incrementer = x => x + 1;

            // Adds 1, then adds 1 to the result
            Func<int, int> doubleIncrementer = incrementer.ApplyTwice();

            Console.WriteLine(doubleIncrementer(5)); // Prints 7

            // Squares the given input
            Func<int, int> squarer = x => x * x;

            // Raises the given input to the power of 4 (squaring twice)
            Func<int, int> pow4 = squarer.ApplyTwice();

            Console.WriteLine(pow4(3)); // Prints 81
        }
    }
}
