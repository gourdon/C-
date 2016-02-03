using System;
using System.Diagnostics.Contracts;

namespace Chapter15
{
    class ImplicitNullReferenceChecking
    {
        static string WontReturnNull()
        {
            Contract.Ensures(Contract.Result<string>() != null);
            return "Guaranteed not to return null";
        }

        static string MightReturnNull()
        {
            return "A later implementation may return null";
        }

        static void Main()
        {
            string literal = "Obviously not null";
            string wontBeNull = WontReturnNull();
            string mightBeNull = MightReturnNull();
            Console.WriteLine(literal.Length);
            Console.WriteLine(wontBeNull.Length);
            Console.WriteLine(mightBeNull.Length);
        }
    }
}
