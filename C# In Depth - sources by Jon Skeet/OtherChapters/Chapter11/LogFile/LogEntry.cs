using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Chapter11.LogFile
{
    class LogEntry
    {
        /// <summary>
        /// Simple lookup of entry types, used for parsing.
        /// </summary>
        static readonly IDictionary<string, EntryType> EntryTypeLookup = ((EntryType[])Enum.GetValues(typeof(EntryType)))
                                                                                           .ToDictionary(entry => entry.ToString());

        public DateTime Timestamp { get; set; }
        public EntryType Type { get; set; }
        public string Message { get; set; }
    
        char[] Tab = {'\t'};
        public LogEntry(string line)
        {
            string[] parts = line.Split(Tab, 3);
            Timestamp = DateTime.ParseExact(parts[0], "u", CultureInfo.InvariantCulture);
            Type = EntryTypeLookup[parts[1]];
            Message = parts[2];
        }

        public LogEntry()
        {
            Timestamp = DateTime.UtcNow;
        }

        public override string ToString()
        {
            return Timestamp.ToString("u", CultureInfo.InvariantCulture) + "\t" + 
                   Type + "\t" + 
                   Message;
        }

        internal void WriteTo(TextWriter writer)
        {
            writer.WriteLine(ToString());
        }
    }
}
