using System.ComponentModel;
using System.IO;

namespace Chapter10
{
    [Description("Listing 10.03")]
    public static class StreamUtilWithExtensions
    {
        const int BufferSize = 8192;

        public static void CopyTo(this Stream input,
                                  Stream output)
        {
            byte[] buffer = new byte[BufferSize];
            int read;
            while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, read);
            }
        }

        public static byte[] ReadFully(this Stream input)
        {
            using (MemoryStream tempStream = new MemoryStream())
            {
                CopyTo(input, tempStream);
                if (tempStream.Length == tempStream.GetBuffer().Length)
                {
                    return tempStream.GetBuffer();
                }
                return tempStream.ToArray();
            }
        }
    }

}
