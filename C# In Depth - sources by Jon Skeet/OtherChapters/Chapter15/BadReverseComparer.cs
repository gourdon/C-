using System.Collections.Generic;
using System.Diagnostics.Contracts;

namespace Chapter15
{
    public class BadReverseComparer<T> : IComparer<T>
    {
        private readonly IComparer<T> original;

        public BadReverseComparer(IComparer<T> original)
        {
            Contract.Requires(original != null);
            this.original = original;
        }

        public int Compare(T x, T y)
        {
            return -original.Compare(x, y);
        }
    }
}
