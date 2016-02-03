using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace Chapter06
{
    /// <summary>
    /// Implementation of a circular buffer; this can be used as a double-ended queue (a deque).
    /// Note that this is not a generic version (although that would be straightforward to write:
    /// this is designed to show how the code in IterationSample might be used.
    /// This code is currently untested! It also doesn't detect if the collection is modified
    /// while it is being iterated over.
    /// </summary>
    public class CircularBuffer : IEnumerable
    {
        private object[] values = new object[0];
        /// <summary>
        /// Index of first entry.
        /// </summary>
        private int start = 0;
        /// <summary>
        /// Number of items in the collection.
        /// </summary>
        private int count = 0;

        /// <summary>
        /// Returns the number of items in the collection.
        /// </summary>
        public int Count
        {
            get { return count; }
        }

        public object this[int index]
        {
            get
            {
                if (index < 0 || index >= count)
                {
                    throw new IndexOutOfRangeException();
                }
                return values[(index + start) % values.Length];
            }
            set
            {
                if (index < 0 || index >= count)
                {
                    throw new IndexOutOfRangeException();
                }
                values[(index + start) % values.Length] = value;
            }
        }

        /// <summary>
        /// Adds a new value at the start of the collection; effectively this
        /// moves the start position back one.
        /// </summary>
        public void AddHead(object value)
        {
            EnsureCapacity(count + 1);
            start--;
            if (start == -1)
            {
                start = values.Length - 1;
            }
            values[start] = value;
            count++;
        }

        /// <summary>
        /// Removes the object at the head of the collection, returning it.
        /// </summary>
        public object RemoveHead()
        {
            if (count == 0)
            {
                throw new InvalidOperationException("Collection was empty");
            }
            object ret = values[start];
            values[start] = null; // Don't prevent previous value from being GC'd
            start = (start + 1) % values.Length;
            count--;
            return ret;
        }

        /// <summary>
        /// Adds a new value at the start of the collection; effectively this
        /// moves the start position back one.
        /// </summary>
        public void AddTail(object value)
        {
            EnsureCapacity(count + 1);
            count++;
            values[count - 1] = value;
        }

        /// <summary>
        /// Removes the object at the tail of the collection, returning it.
        /// </summary>
        public object RemoveTail()
        {
            if (count == 0)
            {
                throw new InvalidOperationException("Collection was empty");
            }
            object ret = this[count - 1];
            values[count - 1] = null; // Don't prevent previous value from being GC'd
            count--;
            return ret;
        }

        /// <summary>
        /// Increases the capacity of the collection if necessary.
        /// </summary>
        public void EnsureCapacity(int newCapacity)
        {
            if (newCapacity <= count)
            {
                return;
            }
            newCapacity = Math.Max(newCapacity, values.Length * 2);
            object[] newValues = new object[newCapacity];
            int countBeforeEndOfArray = Math.Min(values.Length - start, count);
            Array.Copy(values, start, newValues, 0, countBeforeEndOfArray);
            Array.Copy(values, 0, newValues, countBeforeEndOfArray, Math.Max(countBeforeEndOfArray, 0));
            values = newValues;
            start = 0;
        }

        public IEnumerator GetEnumerator()
        {
            return new BufferIterator(this);
        }

        /// <summary>
        /// Note that this is somewhat simpler than IterationSample, as most of the logic is already
        /// in the containing class.
        /// </summary>
        private class BufferIterator : IEnumerator
        {
            CircularBuffer parent;
            int position;

            internal BufferIterator(CircularBuffer parent)
            {
                this.parent = parent;
                position = -1;
            }

            public bool MoveNext()
            {
                if (position != parent.Count)
                {
                    position++;
                }
                return position < parent.Count;
            }

            public object Current
            {
                get
                {
                    if (position < 0 || position >= parent.Count)
                    {
                        throw new InvalidOperationException();
                    }
                    return parent[position];
                }
            }

            public void Reset()
            {
                position = -1;
            }
        }
    }
}
