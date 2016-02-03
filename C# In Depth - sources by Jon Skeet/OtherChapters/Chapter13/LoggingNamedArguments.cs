using System;
using System.ComponentModel;

namespace Chapter13
{
    [Description("Listing 13.04")]
    class LoggingNamedArguments
    {
        static void Dump(int x, int y, int z)
        {
            Console.WriteLine("x={0} y={1} z={2}", x, y, z);
        }

        static int Log(int value)
        {
            Console.WriteLine("Log: {0}", value);
            return value;
        }

        static void Main()
        {
            Dump(x: Log(1), y: Log(2), z: Log(3));
            Dump(z: Log(3), x: Log(1), y: Log(2));
        }
    }
}
