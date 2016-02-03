using System;
using System.ComponentModel;

namespace Chapter13
{
    [Description("Listing 13.01")]
    class SimpleOptionalParameters
    {
        static void Dump(int x, int y = 20, int z = 30)
        {
            Console.WriteLine("x={0} y={1} z={2}", x, y, z);
        }
        
        static void Main()
        {
            Dump(1, 2, 3);
            Dump(1, 2);
            Dump(1);
        }
    }
}
