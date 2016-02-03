using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter08
{
    [Description("Listing 8.2")]
    public class Person
    {
        public int Age { get; set; }
        public string Name { get; set; }

        List<Person> friends = new List<Person>();
        public List<Person> Friends
        {
            get { return friends; }
        }

        Location home = new Location();
        public Location Home
        {
            get { return home; }
        }

        public Person()
        {
        }

        public Person(string name)
        {
            Name = name;
        }
    }
}
