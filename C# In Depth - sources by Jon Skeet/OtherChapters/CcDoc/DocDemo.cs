using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics.Contracts;

namespace CcDoc
{
    /// <summary>
    /// Class summary.
    /// </summary>
    public sealed class DocDemo
    {
        private int callCount = 0;

        [ContractInvariantMethod]
        private void Invariant()
        {
            Contract.Invariant(callCount >= 0);
            Contract.Invariant(callCount < 100, "Wrap at 100.");
        }

        /// <summary>
        /// Method summary.
        /// </summary>
        /// <returns>The input, reversed.</returns>
        public string Reverse(string text)
        {
            Contract.Requires<ArgumentNullException>
                (text != null, "text");
            Contract.Ensures(Contract.Result<string>() != null);

            callCount = (callCount + 1) % 100;

            char[] chars = text.ToCharArray();
            Array.Reverse(chars);
            return new string(chars);
        }
    }
}
