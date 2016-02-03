using System.ComponentModel;

namespace Chapter08
{
    [Description("Listing 8.1")]
    public class InstanceCountingPerson
    {
        public string Name { get; private set; }
        public int Age { get; private set; }

        private static int InstanceCounter { get; set; }
        private static readonly object counterLock = new object();

        public InstanceCountingPerson(string name, int age)
        {
            Name = name;
            Age = age;

            lock (counterLock)
            {
                InstanceCounter++;
            }
        }
    }
}
