using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics.Contracts;

[assembly:ContractVerification(false)]

namespace SelectiveCheckingDemo
{
    [ContractVerification(true)]
    public class CheckedType
    {
        public void CheckedMethod() {}

        [ContractVerification(false)]
        public void UncheckedMethod() {}
    }

    public class UncheckedType
    {
        public void UncheckedMethod() {}

        [ContractVerification(true)]
        public void CheckedMethod() {}
    }
}
