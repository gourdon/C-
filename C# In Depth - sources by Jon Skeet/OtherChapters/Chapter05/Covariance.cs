using System;
using System.ComponentModel;
using System.IO;

namespace Chapter05
{
    [Description("Listing 5.03")]
    class Covariance
    {
        delegate Stream StreamFactory();

        static MemoryStream GenerateSampleData()
        {
            byte[] buffer = new byte[16];
            for (int i = 0; i < buffer.Length; i++)
            {
                buffer[i] = (byte)i;
            }
            return new MemoryStream(buffer);
        }

        static void Main()
        {
            StreamFactory factory = GenerateSampleData;

            using (Stream stream = factory())
            {
                int data;
                while ((data = stream.ReadByte()) != -1)
                {
                    Console.WriteLine(data);
                }
            }
        }
    }
}
