using System;
using System.ComponentModel;

namespace Chapter04
{
    [Description("Listing 4.2")]
    class BoxingAndUnboxing
    {
        static void Main()
        {
            Nullable<int> nullable = 5;

            object boxed = nullable;
            Console.WriteLine(boxed.GetType());

            int normal = (int)boxed;
            Console.WriteLine(normal);

            nullable = (Nullable<int>)boxed;
            Console.WriteLine(nullable);

            nullable = new Nullable<int>();
            boxed = nullable;
            Console.WriteLine(boxed == null);

            nullable = (Nullable<int>)boxed;
            Console.WriteLine(nullable.HasValue);
        }
    }
}
