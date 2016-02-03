using System;

namespace Chapter09.HigherOrderFunctions
{
    static class CurryingExtensions
    {
        public static Func<A, Func<B, R>> Curry<A, B, R>(this Func<A, B, R> f)
        {
            // Return a function which, given a, will return a function which,
            // given b, will return f(a, b) where f is the parameter provided above
            return a => b => f(a, b);
        }
    }

    // See http://blogs.msdn.com/wesdyer/archive/2007/01/29/currying-and-partial-function-application.aspx
    class PartialApplication
    {

        static void Main()
        {
            // Start with a simple adding delegate
            Func<int, int, int> adder = (x, y) => (x + y);

            // Now curry it
            Func<int, Func<int, int>> curried = adder.Curry();

            Func<int, int> addsTwo = curried(2);

            Console.WriteLine(addsTwo(5)); // Prints 7

            Func<int, int> addsTen = curried(10);
            Console.WriteLine(addsTen(3)); // Prints 13

            // Can I lie down now please?
        }
    }
}
