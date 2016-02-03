using System;
using System.ComponentModel;

namespace Chapter08
{
    [Description("Listing 8.4")]
    class FirstAnonymousType
    {
        static void Main()
        {
            var tom = new { Name = "Tom", Age = 6 };
            var holly = new { Name = "Holly", Age = 34 };
            var jon = new { Name = "Jon", Age = 33 };

            Console.WriteLine("{0} is {1} years old", jon.Name, jon.Age);
        }
    }
}
