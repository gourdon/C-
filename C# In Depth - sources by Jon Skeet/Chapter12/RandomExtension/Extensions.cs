using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;

namespace RandomExtension
{
    [Description("Listing 12.17")]
    public static class Extensions
    {
        /// <summary>
        /// Returns a random element from the given source.
        /// </summary>
        /// <remarks>
        /// <para>
        /// If the source sequence implements <see cref="IList{T}"/>, this method
        /// will pick a random index based on the count of the list, and then
        /// return that element directly.
        /// </para>
        /// <para>
        /// If the source sequence implements <see cref="ICollection{T}"/> but not
        /// <see cref="IList{T}"/>, the index will still be chosen randomly based
        /// on the count, but then the element at that index will be returned
        /// by iterating that many times.
        /// </para>
        /// <para>
        /// Otherwise, a random replacement strategy is used: we start with the
        /// first element as our "current" value, and replace it with the second
        /// element with a probability of 1/2. The third element replaces "current"
        /// with a probability of 1/3, and so on until we reach the end of the sequence.
        /// This will yield a uniform distribution, but requires as many calls to
        /// the random number generator as there are elements in the sequence - and
        /// the entire sequence will be iterated over before returning. The sequence
        /// will only be evaluated once, however, so this method can still be used
        /// with long sequences which cannot be replayed.
        /// </para>
        /// </remarks>
        /// <typeparam name="T">The type of the elements in source.</typeparam>
        /// <param name="source">A sequence of values to pick a random element from</param>
        /// <param name="random">A random number generator</param>
        /// <returns>An element picked at random from the source</returns>
        /// <exception cref="ArgumentNullException">
        /// <paramref name="source"/> or <paramref name="random"/> is null.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// <paramref name="source"/> is empty.
        /// </exception>
        public static T RandomElement<T>(this IEnumerable<T> source,
                                         Random random)
        {
            if (source == null)
            {
                throw new ArgumentNullException("source");
            }
            if (random == null)
            {
                throw new ArgumentNullException("random");
            }
            ICollection collection = source as ICollection;
            if (collection != null)
            {
                int count = collection.Count;
                if (count == 0)
                {
                    throw new InvalidOperationException("Sequence was empty.");
                }
                int index = random.Next(count);
                return source.ElementAt(index);
            }
            ICollection<T> genericCollection = source as ICollection<T>;
            if (genericCollection != null)
            {
                int count = genericCollection.Count;
                if (count == 0)
                {
                    throw new InvalidOperationException("Sequence was empty.");
                }
                int index = random.Next(count);
                return source.ElementAt(index);
            }
            using (IEnumerator<T> iterator = source.GetEnumerator())
            {
                if (!iterator.MoveNext())
                {
                    throw new InvalidOperationException("Sequence was empty.");
                }
                int countSoFar = 1;
                T current = iterator.Current;
                while (iterator.MoveNext())
                {
                    countSoFar++;
                    if (random.Next(countSoFar) == 0)
                    {
                        current = iterator.Current;
                    }
                }
                return current;
            }
        }
    }
}
