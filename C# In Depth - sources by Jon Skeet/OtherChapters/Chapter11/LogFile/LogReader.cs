using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Chapter11.LogFile
{
    class LogReader
    {
        public static IEnumerable<string> ReadLines(string filename)
        {
            using (TextReader reader = File.OpenText(filename))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    yield return line;
                }
            }
        }

        static void Main()
        {
            var query = from file in Directory.GetFiles(@"C:\CSharpInDepthLogs", "*.log")
                        from line in ReadLines(file)
                        let entry = new LogEntry(line)
                        where entry.Type == EntryType.Error
                        select entry;

            Console.WriteLine("Total errors: {0}", query.Count());
        }
    }
}
