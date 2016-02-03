using System;
using System.Diagnostics.Contracts;

namespace Chapter15
{
    class MaskingContractFailure
    {
        static void RequireNonNullArgument(string text)
        {
            Contract.Requires(text != null, "Don't pass in null");
            Console.WriteLine("In method body");
        }

        static void HandleFailure(object sender, ContractFailedEventArgs args)
        {
            Console.WriteLine("{0}: {1} {2}", args.FailureKind,
                args.Condition, args.Message);
            args.SetHandled();
        }

        static void Main()
        {
            Contract.ContractFailed += HandleFailure;
            RequireNonNullArgument(null);
        }
    }
}
