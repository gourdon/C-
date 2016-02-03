using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics.Contracts;

namespace ContractAssemblyDemo
{
    public class PreconditionDemo
    {
        public void DontPassInNull(string text)
        {
            Contract.Requires(text != null);
            Console.WriteLine("In DontPassInNull()");
        }
    }
}
