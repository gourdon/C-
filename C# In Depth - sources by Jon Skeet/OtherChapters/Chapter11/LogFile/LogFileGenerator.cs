using System;
using System.IO;
using Chapter11.Queries;
using MiscUtil;

namespace Chapter11.LogFile
{
    /// <summary>
    /// Generates large log files for use in other programs
    /// </summary>
    class LogFileGenerator
    {
        const int FilesToGenerate = 20;
        const int EntriesPerFile = 100000;
        static readonly long TicksPerEntry = TimeSpan.FromDays(1).Ticks / EntriesPerFile;
        static readonly int NumberOfEntryTypes = Enum.GetValues(typeof(EntryType)).Length;

        static void Main()
        {
            DirectoryInfo directory = new DirectoryInfo(@"c:\CSharpInDepthLogs");
            if (directory.Exists)
            {
                directory.Delete(true);
            }
            directory.Create();

            DateTime today = DateTime.UtcNow.Date;

            DateTime start = DateTime.Now;
            foreach (DateTime date in new DateTimeRange(today.AddDays(-FilesToGenerate),
                                                        today.AddDays(-1)))
            {
                using (TextWriter writer = new StreamWriter(Path.Combine(directory.FullName, date.ToString("yyyyMMdd'.log'"))))
                {
                    for (int i = 0; i < EntriesPerFile; i++)
                    {
                        LogEntry entry = new LogEntry
                        {
                            Timestamp = date + TimeSpan.FromTicks(TicksPerEntry*i),
                            Type = (EntryType)StaticRandom.Next(NumberOfEntryTypes),
                            Message = GenerateRandomMessage()
                        };
                        entry.WriteTo(writer);
                    }
                }
            }
            DateTime end = DateTime.Now;
            Console.WriteLine(end - start);
        }

        static string GenerateRandomMessage()
        {
            return "This is a random message, with a random number: " + StaticRandom.Next(1000);
        }
    }
}
