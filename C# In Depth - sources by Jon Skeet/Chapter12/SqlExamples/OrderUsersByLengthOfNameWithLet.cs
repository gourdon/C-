using System;
using System.ComponentModel;
using System.Linq;

using Model;

namespace SqlExamples
{
    [Description("Listing 12.2")]
    class OrderUsersByLengthOfNameWithLet
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                var query = from user in context.Users
                            let length = user.Name.Length
                            orderby length
                            select new { Name = user.Name, Length = length };

                foreach (var entry in query)
                {
                    Console.WriteLine("{0}: {1}", entry.Length, entry.Name);
                }
            }
        }
    }
}
