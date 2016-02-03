using System;
using System.ComponentModel;

namespace Chapter13
{
    [Description("Listing 13.05")]
    class EvilSideEffects
    {
        static void Dump(int x, int y, int z)
        {
            Console.WriteLine("x={0} y={1} z={2}", x, y, z);
        }

        static void Main()
        {
            int i = 0;
            Dump(x: ++i, y: ++i, z: ++i);
            i = 0;
            Dump(z: ++i, x: ++i, y: ++i);
        }
    }
}
