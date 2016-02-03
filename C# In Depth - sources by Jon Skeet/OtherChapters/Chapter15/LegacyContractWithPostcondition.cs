using System;
using System.Diagnostics.Contracts;
using System.Linq;

namespace Chapter15
{
    class LegacyContractWithPostCondition
    {
        static int CountWhitespace(string text)
        {
            if (text == null)
            {
                throw new ArgumentNullException("text");
            }
            Contract.Ensures(Contract.Result<int>() > 0);
            return text.Count(char.IsWhiteSpace);
        }

        static void Main()
        {
            Console.WriteLine(CountWhitespace("x y z"));
        }
    }
}
