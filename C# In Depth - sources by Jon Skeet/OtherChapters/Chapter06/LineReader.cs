using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;

namespace Chapter06
{
    [Description("Listing 6.8")]
    class LineReader
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
            foreach (string line in ReadLines("../../LineReader.cs"))
            {
                Console.WriteLine(line);
            }
        }
    }
}
