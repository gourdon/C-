using System;
using System.ComponentModel;

namespace Chapter14
{
    [Description("Listing 14.19")]
    class InheritanceOverloading
    {
        class Base
        {
            public void Execute(object x)
            {
                Console.WriteLine("object");
            }
        }

        class Derived : Base
        {
            public void Execute(string x)
            {
                Console.WriteLine("string");
            }
        }

        static void Main()
        {
            Base receiver = new Derived();
            dynamic d = "text";
            receiver.Execute(d);
        }
    }
}
