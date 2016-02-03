using System;
using System.ComponentModel;

namespace Chapter11.Queries
{
    static class Extensions
    {
        public static Dummy<T> Where<T>(this Dummy<T> dummy,
                                        Func<T,bool> predicate)
        {
            Console.WriteLine ("Where called");
            return dummy;
        }
    }

    class Dummy<T>
    {
        public Dummy<U> Select<U>(Func<T,U> selector)
        {
            Console.WriteLine ("Select called");
            return new Dummy<U>();
        }
    }

    [Description("Listing 11.03")]
    class TranslationExample
    {
        static void Main()
        {
            var source = new Dummy<string>();

            var query = from dummy in source
                        where dummy.ToString()=="Ignored"
                        select "Anything";
        }
    }
}

