using System;
using System.Diagnostics.Contracts;
using System.Linq;

namespace Chapter15
{
    class CodeContractsWhitespaceCountWithArgumentNullException
    {
        /// <summary>Counts the number of whitespace characters
        /// in <paramref name="text"/>.</summary>
        /// <param name="text">String to examine. Must not be null.</param>
        /// <returns>The number of whitespace characters.</returns>
        static int CountWhitespace(string text)
        {
            Contract.Requires<ArgumentNullException>(text != null, "text");
            return text.Count(char.IsWhiteSpace);
        }

        static void Main()
        {
            Console.WriteLine(CountWhitespace("x y z"));
        }
    }
}
