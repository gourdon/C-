using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics.Contracts;

namespace ContractAssemblyDemo
{
    public class PostconditionDemo
    {
        public int GimmeFive()
        {
            Contract.Ensures(Contract.Result<int>() == 5);
            Console.WriteLine("In GimmeFive()");
            return 5;
        }
    }
}
