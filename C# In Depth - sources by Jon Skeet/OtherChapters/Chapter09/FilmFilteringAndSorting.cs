using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter09
{
    [Description("Listing 9.04")]
    class FilmFilteringAndSorting
    {
        class Film
        {
            public string Name { get; set; }
            public int Year { get; set; }
        }

        static void Main()
        {
            var films = new List<Film>
            {
                new Film {Name="Jaws", Year=1975},
                new Film {Name="Singing in the Rain", Year=1952},
                new Film {Name="Some like it Hot", Year=1959},
                new Film {Name="The Wizard of Oz", Year=1939},
                new Film {Name="It's a Wonderful Life", Year=1946},
                new Film {Name="American Beauty", Year=1999},
                new Film {Name="High Fidelity", Year=2000},
                new Film {Name="The Usual Suspects", Year=1995}
            };

            Action<Film> print = 
                film => Console.WriteLine("Name={0}, Year={1}", film.Name, film.Year);

            // Note: extra lines added for clarity when running
            Console.WriteLine("All films");
            films.ForEach(print);                                
            Console.WriteLine();

            Console.WriteLine("Oldies");
            films.FindAll(film => film.Year < 1960)              
                 .ForEach(print);
            Console.WriteLine();

            Console.WriteLine("Sorted");
            films.Sort((f1, f2) => f1.Name.CompareTo(f2.Name));  
            films.ForEach(print);                                
        }        
    }
}
