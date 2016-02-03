using System;
using System.Diagnostics.Contracts;

namespace Chapter15
{
    class NullArgumentChecking
    {
        static string GoBangIfNull(string message)
        {
            Contract.Requires(message != null);
            Contract.Ensures(Contract.Result<string>() != null);
            Console.WriteLine(message);
            return message;
        }

        static void Main()
        {
            GoBangIfNull(null);
        }
    }
}
