using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter03
{
    [Description("Listing 3.06")]
    public sealed class Pair<TFirst, TSecond> : IEquatable<Pair<TFirst, TSecond>>
    {
        private static readonly IEqualityComparer<TFirst> FirstComparer =
           EqualityComparer<TFirst>.Default;
        private static readonly IEqualityComparer<TSecond> SecondComparer =
           EqualityComparer<TSecond>.Default;

        private readonly TFirst first;
        private readonly TSecond second;

        public Pair(TFirst first, TSecond second)
        {
            this.first = first;
            this.second = second;
        }

        public TFirst First { get { return first; } }

        public TSecond Second { get { return second; } }

        public bool Equals(Pair<TFirst, TSecond> other)
        {
            return other != null &&
               FirstComparer.Equals(this.First, other.First) &&
               SecondComparer.Equals(this.Second, other.Second);
        }

        public override bool Equals(object o)
        {
            return Equals(o as Pair<TFirst, TSecond>);
        }

        public override int GetHashCode()
        {
            return FirstComparer.GetHashCode(first) * 37 +
               SecondComparer.GetHashCode(second);
        }
    }

    [Description("Listing 3.07")]
    public static class Pair
    {
        public static Pair<TFirst, TSecond> Of<TFirst, TSecond>(TFirst first, TSecond second)
        {
            return new Pair<TFirst, TSecond>(first, second);
        }
    }
}