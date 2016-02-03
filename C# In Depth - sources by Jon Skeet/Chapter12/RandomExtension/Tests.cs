using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;

namespace RandomExtension
{
    /// <summary>
    /// Note to readers: testing randomness is quite tricky. We
    /// don't want to specify in the contract *how* it should use
    /// the random number generator. This means these tests have
    /// somewhat less strictness about what to expect than normal
    /// unit tests.
    /// </summary>
    [TestFixture]
    public class Tests
    {
        private static readonly IEnumerable<int> Sample = Enumerable.Range(0, 100);

        [Test]
        public void NullSource_ThrowsArgumentNullException()
        {
            IEnumerable<int> source = null;
            Assert.Throws<ArgumentNullException>(() => source.RandomElement(new Random()));
        }

        [Test]
        public void NullRandom_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => Sample.RandomElement(null));
        }

        [Test]
        public void EmptyListSource_ThrowsInvalidOperationException()
        {
            IEnumerable<int> source = new List<int>();
            Assert.Throws<InvalidOperationException>(() => source.RandomElement(new Random()));
        }

        [Test]
        public void EmptyLazySource_ThrowsInvalidOperationException()
        {
            IEnumerable<int> source = Enumerable.Range(0, 0);
            Assert.Throws<InvalidOperationException>(() => source.RandomElement(new Random()));
        }

        [Test]
        public void ListSource_GivesReasonableDistribution()
        {
            VerifyDistribution(Sample.ToList());
        }

        [Test]
        public void CollectionSource_GivesReasonableDistribution()
        {
            VerifyDistribution(new LinkedList<int>(Sample));
        }

        /// <summary>
        /// HashSet[T] implements ICollection[T] but not ICollection; we're relying
        /// on its order staying the same after we've finished populating it, but we
        /// don't actually care about what that order is.
        /// </summary>
        [Test]
        public void GenericCollectionSource_GivesReasonableDistribution()
        {
            VerifyDistribution(new HashSet<int>(Sample));
        }

        [Test]
        public void LazySource_GivesReasonableDistribution()
        {
            VerifyDistribution(Sample);
        }

        [Test]
        public void VerifyDistribution(IEnumerable<int> source)
        {
            Random random = new Random();
            // Take 100,000 samples... each value should come up *roughly* 1000 times
            var groups = Enumerable.Range(0, 100000)
                                   .Select(x => source.RandomElement(random))
                                   .GroupBy(x => x)
                                   .ToList(); // Only do it once!
            // First check there actually *were* 100 groups (i.e. every value came up at least once)
            Assert.AreEqual(100, groups.Count());
            foreach (var group in groups)
            {
                Assert.That(group.Count(), Is.InRange(800, 1200), "Count for " + group.Key);
            }
        }
    }
}
