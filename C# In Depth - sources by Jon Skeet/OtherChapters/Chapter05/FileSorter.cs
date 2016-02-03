using System;
using System.ComponentModel;
using System.IO;

namespace Chapter05
{
    [Description("Listing 5.08")]
    class FileSorter
    {
        static void SortAndShowFiles(string title,
                                     Comparison<FileInfo> sortOrder)
        {
            FileInfo[] files = new DirectoryInfo(@"C:\").GetFiles();

            Array.Sort(files, sortOrder);

            Console.WriteLine(title);
            foreach (FileInfo file in files)
            {
                Console.WriteLine("  {0} ({1} bytes)",
                                   file.Name, file.Length);
            }
        }

        static void Main()
        {
            SortAndShowFiles("Sorted by name:", delegate(FileInfo first, FileInfo second)
                { return first.Name.CompareTo(second.Name); }
            );


            SortAndShowFiles("Sorted by length:", delegate(FileInfo first, FileInfo second)
                { return first.Length.CompareTo(second.Length); }
            );            
        }
    }
}
