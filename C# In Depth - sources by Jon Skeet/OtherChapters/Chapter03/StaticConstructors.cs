using System;
using System.ComponentModel;

namespace Chapter03
{
    [Description("Listing 3.09")]
    class StaticConstructors
    {
        class Outer<T>
        {
            public class Inner<U, V>
            {
                static Inner()
                {
                    Console.WriteLine("Outer<{0}>.Inner<{1},{2}>",
                                      typeof(T).Name,
                                      typeof(U).Name,
                                      typeof(V).Name);
                }

                public static void DummyMethod()
                {
                }
            }
        }

        static void Main()
        {
            Outer<int>.Inner<string, DateTime>.DummyMethod();
            Outer<string>.Inner<int, int>.DummyMethod();
            Outer<object>.Inner<string, object>.DummyMethod();
            Outer<string>.Inner<string, object>.DummyMethod();
            Outer<object>.Inner<object, string>.DummyMethod();
            Outer<string>.Inner<int, int>.DummyMethod();
        }
    }
}
