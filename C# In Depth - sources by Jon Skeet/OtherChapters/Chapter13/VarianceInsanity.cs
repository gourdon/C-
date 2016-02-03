using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Chapter13
{
    class VarianceInsanity
    {
        delegate Func<T> FuncFunc<out T>();
        delegate void ActionAction<out T>(Action<T> action);
        delegate void FuncAction<in T>(Func<T> function);
        delegate Action<T> ActionFunc<in T>();
    }
}
