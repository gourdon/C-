using System.Diagnostics.Contracts;

namespace Chapter15
{
    class BreakingInvariantInheritance
    {
        abstract class Base
        {
            [ContractInvariantMethod]
            private void BaseInvariant()
            {
                Contract.Invariant(true);
            }

            public void AttemptToBreakDerivedInvariant()
            {
                BreakDerivedInvariant();
                // What is called here? Just BaseInvariant,
                // or DerivedInvariant too?
            }

            protected abstract void BreakDerivedInvariant();
        }

        sealed class Derived : Base
        {
            private bool valid = true;

            [ContractInvariantMethod]
            private void DerivedInvariant()
            {
                Contract.Invariant(valid);
            }

            protected override void BreakDerivedInvariant()
            {
                valid = false;
            }
        }

        static void Main()
        {
            Derived x = new Derived();
            x.AttemptToBreakDerivedInvariant();
        }
    }
}
