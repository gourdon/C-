using System.Collections.Generic;

namespace Chapter03.Variance
{
    public sealed class ComparisonHelper<TBase, TDerived> : IComparer<TDerived>
        where TDerived : TBase
    {
        private readonly IComparer<TBase> comparer;

        public ComparisonHelper(IComparer<TBase> comparer)
        {
            this.comparer = comparer;
        }

        public int Compare(TDerived x, TDerived y)
        {
            // Use the implicit conversions of x and y to TBase
            return comparer.Compare(x, y);
        }
    }
}