using System;
using System.ComponentModel;

namespace Chapter08
{
    [Description("Listing 8.5")]
    class AnonymousTypeInArray
    {
        static void Main()
        {
            var family = new[]                         
            {
                new { Name = "Holly", Age = 34 },      
                new { Name = "Jon", Age = 33 },        
                new { Name = "Tom", Age = 6 },         
                new { Name = "Robin", Age = 3 },       
                new { Name = "William", Age = 3 }      
            };

            int totalAge = 0;
            foreach (var person in family)
            {
                totalAge += person.Age;
            }
            Console.WriteLine("Total age: {0}", totalAge);
        }
    }
}
