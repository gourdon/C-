using System;

namespace Chapter07
{
    partial class Example<TFirst, TSecond>
        : EventArgs, IDisposable
    {
        public void Dispose()
        {
        }
    }
}