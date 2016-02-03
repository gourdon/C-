using System;
using System.ComponentModel;

namespace Chapter13
{
    [Description("Listing 13.03")]
    class SimpleNamedArguments
    {
        static void Dump(int x, int y, int z)
        {
            Console.WriteLine("x={0} y={1} z={2}", x, y, z);
        }

        static void Main()
        {
            Dump(1, 2, 3);
            Dump(x: 1, y: 2, z: 3);
            Dump(z: 3, y: 2, x: 1);
            Dump(1, y: 2, z: 3);
            Dump(1, z: 3, y: 2);
        }
    }
}
