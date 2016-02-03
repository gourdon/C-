using System;
using System.ComponentModel;

namespace Chapter09
{
    [Description("Listing 9.13")]
    class ReturnTypeInferenceWithMultipleReturns
    {
        delegate T MyFunc<T>();

        static void WriteResult<T>(MyFunc<T> function)
        {
            Console.WriteLine(function());
        }

        static void Main()
        {
            WriteResult(delegate
            {
                if (DateTime.Now.Hour < 12)
                {
                    return 10;
                }
                else
                {
                    return new object();
                }
            });
        }
    }
}
