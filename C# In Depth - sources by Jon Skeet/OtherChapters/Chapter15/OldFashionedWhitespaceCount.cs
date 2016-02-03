using System;
using System.Linq;

namespace Chapter15
{
    class OldFashionedWhitespaceCount
    {
        /// <summary>Counts the number of whitespace characters
        /// in <paramref name="text"/>.</summary>
        /// <param name="text">String to examine. Must not be null.</param>
        /// <exception cref="ArgumentNullException">
        /// <paramref name="text"/> is null.</exception>
        /// <returns>The number of whitespace characters.</returns>
        static int CountWhitespace(string text)
        {
            if (text == null)
            {
                throw new ArgumentNullException("text");
            }
            return text.Count(char.IsWhiteSpace);
        }

        static void Main()
        {
            Console.WriteLine(CountWhitespace("x y z"));
        }
    }
}
