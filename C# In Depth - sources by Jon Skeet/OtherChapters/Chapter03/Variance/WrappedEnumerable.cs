using System.Collections;
using System.Collections.Generic;

namespace Chapter03.Variance
{
    public sealed class WrappedEnumerable<TOriginal, TWrapper> : IEnumerable<TWrapper>
       where TOriginal : TWrapper
    {
        private readonly IEnumerable<TOriginal> originals;

        public WrappedEnumerable(IEnumerable<TOriginal> originals)
        {
            this.originals = originals;
        }

        public IEnumerator<TWrapper> GetEnumerator()
        {
            foreach (TOriginal item in originals)
            {
                yield return item;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
