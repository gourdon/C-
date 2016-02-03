using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter03
{
    [Description("Listing 3.14")]
    class ComparisonHelper<TBase, TDerived> : IComparer<TDerived>
        where TDerived : TBase
    {
        private readonly IComparer<TBase> comparer;

        public ComparisonHelper(IComparer<TBase> comparer)
        {
            this.comparer = comparer;
        }

        public int Compare(TDerived x, TDerived y)
        {
            return comparer.Compare(x, y);
        }
    }
}
