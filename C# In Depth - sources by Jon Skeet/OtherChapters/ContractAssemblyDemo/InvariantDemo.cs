using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics.Contracts;

namespace ContractAssemblyDemo
{
    public class InvariantDemo
    {
        private int alwaysZero;

        public void DoNothing()
        {
            Console.WriteLine("In DoNothing()");
        }

        [ContractInvariantMethod]
        private void ObjectInvariant()
        {
            Contract.Invariant(alwaysZero == 0);
        }
    }
}
