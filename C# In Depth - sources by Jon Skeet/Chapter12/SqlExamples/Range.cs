using System;
using System.Collections;
using System.Collections.Generic;

namespace SqlExamples
{
    public abstract class Range<T> : IEnumerable<T>
        where T : IComparable<T>
    {
        T start;
        T end;

        public Range(T start, T end)
        {
            if (start.CompareTo(end) > 0)
            {
                throw new ArgumentOutOfRangeException();
            }
            this.start = start;
            this.end = end;
        }

        public T Start
        {
            get { return start; }
        }

        public T End
        {
            get { return end; }
        }

        public bool Contains(T value)
        {
            return value.CompareTo(start) >= 0 &&
                   value.CompareTo(end) <= 0;
        }

        public IEnumerator<T> GetEnumerator()
        {
            T value = start;
            while (value.CompareTo(end) <= 0)
            {
                yield return value;
                value = GetNextValue(value);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        protected abstract T GetNextValue(T current);
    }
}