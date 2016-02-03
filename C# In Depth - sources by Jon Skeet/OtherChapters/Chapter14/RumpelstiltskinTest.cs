using System;
using System.ComponentModel;

namespace Chapter14
{
    [Description("Listing 14.34")]
    class RumpelstiltskinTest
    {
        static void Main()
        {
            Console.WriteLine("Testing first implementation");
            dynamic x = new Rumpelstiltskin("Hermione");
            x.Harry();
            x.Ron();
            x.Hermione();

            Console.WriteLine("Testing alternative implementation");
            // Now check the alternative implementation
            x = new Rumpelstiltskin("Ron");
            x.Harry();
            x.Ron();
            x.Hermione();
        }
    }
}
