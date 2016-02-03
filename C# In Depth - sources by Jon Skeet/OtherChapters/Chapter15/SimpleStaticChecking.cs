using System.Diagnostics.Contracts;

namespace Chapter15
{
    class SimpleStaticChecking
    {
        static string DontPassMeNull(string input)
        {
            Contract.Requires(input != null);
            Contract.Ensures(Contract.Result<string>() != null);
            return input;
        }

        static string MightReturnNull()
        {
            return "Not null really";
        }

        static void Main(string[] args)
        {
            DontPassMeNull("definitely okay");
            DontPassMeNull(MightReturnNull());
            DontPassMeNull(null);            
        }
    }
}
