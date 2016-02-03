using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Chapter11.Queries
{
    /// <summary>
    /// The listings are scattered around .cs files within various directories.
    /// This class uses LINQ to find all classes with a suitable Description
    /// attribute, groups them by chapters and orders them by chapter and listing number.
    /// </summary>
    class DisplayListingsMap
    {
        static readonly Regex ListingPattern = new Regex(
            @"# First match the start of the attribute, up to the bit we're interested in
            \[Description\(\""Listing\ 
            # The 'text' group is the whole of the description after Listing
            (?<text>
            # The 'chapter' group is the first set of digits in the description, before a dot
            (?<chapter>\d+)\.
            # The chapter group is the second set of digits in the description
            (?<listing>\d+)
            # After that we don't care - stop the 'text' group at the double quote
            [^\""]*)
            # Now match the end of the attribute
            \""\)\]",
            RegexOptions.Compiled | RegexOptions.IgnorePatternWhitespace);

        static void Main()
        {
            DirectoryInfo directory = new DirectoryInfo(@"..\..\..\..");

            var query = from file in directory.GetFiles("*.cs", SearchOption.AllDirectories)
                        let match = ListingPattern.Match(File.ReadAllText(file.FullName))
                        where match.Success
                        let Details = new
                        {
                            Text = match.Groups["text"].Value,
                            Chapter = int.Parse(match.Groups["chapter"].Value),
                            Listing = int.Parse(match.Groups["listing"].Value)
                        }
                        orderby Details.Chapter, Details.Listing
                        group new { File = file, Description=Details.Text } by Details.Chapter;

            foreach (var chapter in query)
            {
                Console.WriteLine("Chapter {0}", chapter.Key);
                foreach (var listing in chapter)
                {
                    Console.WriteLine("{0}: {1}", listing.Description, listing.File.Name);
                }
                Console.WriteLine();                
            }
        }
    }
}
