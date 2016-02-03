using System;
using System.ComponentModel;

namespace Chapter07
{
    [Description("Listing 7.01")]
    partial class Example<TFirst, TSecond>
        : IEquatable<string>
        where TFirst : class
    {
        public bool Equals(string other)
        {
            return false;
        }
    }
}