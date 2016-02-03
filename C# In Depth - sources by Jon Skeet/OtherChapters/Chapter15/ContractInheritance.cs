using System;
using System.Diagnostics.Contracts;

namespace Chapter15
{
    class ContractInheritance
    {
        [Pure]
        static bool Report(string text)
        {
            Console.WriteLine(text);
            return true;
        }

        class Base
        {
            [ContractInvariantMethod]
            private void BaseInvariant()
            {
                Contract.Invariant(Report("BaseInvariant"));
            }

            public virtual void VirtualMethod(string text)
            {
                Contract.Requires(Report("Base precondition"));
                Contract.Ensures(Report("Base postcondition"));
            }
        }

        class Derived : Base
        {
            [ContractInvariantMethod]
            private void DerivedInvariant()
            {
                Contract.Invariant(Report("DerivedInvariant"));
            }

            public override void VirtualMethod(string text)
            {
                Contract.Ensures(Report("Derived postcondition"));
            }
        }

        class DerivedWithNoInvariant : Base
        {
            public void BaseInvariantWillStillBeCalled()
            {                
            }
        }

        static void Main()
        {
            Base d = new Derived();
            Console.WriteLine("Constructor finished");
            d.VirtualMethod("");
        }
    }
}
