using System;
using System.Diagnostics.Contracts;
using System.Linq;

namespace Chapter15
{
    class LegacyContractsWhitespaceCount
    {
        static int CountWhitespace(string text)
        {
            if (text == null)
            {
                throw new ArgumentNullException("text");
            }
            Contract.EndContractBlock();
            return text.Count(char.IsWhiteSpace);
        }

        static void Main()
        {
            Console.WriteLine(CountWhitespace("x y z"));
        }
    }
}
