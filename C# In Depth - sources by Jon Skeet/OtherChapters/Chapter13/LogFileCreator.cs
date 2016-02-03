using System;
using System.ComponentModel;
using System.IO;
using System.Text;

namespace Chapter13
{
    [Description("Listing 13.02 and 06")]
    class LogFileCreator
    {
        static void AppendTimestamp(string filename,
                                    string message,
                                    Encoding encoding = null,
                                    DateTime? timestamp = null)
        {
            Encoding realEncoding = encoding ?? Encoding.UTF8;
            DateTime realTimestamp = timestamp ?? DateTime.Now;
            using (TextWriter writer = new StreamWriter
                (filename, true, realEncoding))
            {
                writer.WriteLine("{0:s}: {1}", realTimestamp, message);
            }
        }

        static void Main()
        {
            AppendTimestamp("utf8.txt", "First message");
            AppendTimestamp("ascii.txt", "ASCII", Encoding.ASCII);
            AppendTimestamp("utf8.txt", "Message in the future", null,
                            new DateTime(2030, 1, 1));

            // From listing 13.06; no need for encoding argument
            AppendTimestamp("utf8.txt", "Message in the future",
                            timestamp: new DateTime(2030, 1, 1));
        }
    }
}
