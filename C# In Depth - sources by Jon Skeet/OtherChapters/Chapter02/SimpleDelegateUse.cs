using System;
using System.ComponentModel;

namespace Chapter02
{
    delegate void StringProcessor(string input);

    class Person
    {
        string name;

        public Person(string name)
        {
            this.name = name;
        }

        public void Say(string message)
        {
            Console.WriteLine("{0} says: {1}", name, message);
        }
    }

    class Background
    {
        public static void Note(string note)
        {
            Console.WriteLine("({0})", note);
        }
    }

    [Description("Listing 2.1")]
    class SimpleDelegateUse
    {
        static void Main()
        {
            Person jon = new Person("Jon");
            Person tom = new Person("Tom");

            StringProcessor jonsVoice, tomsVoice, background;
            jonsVoice = new StringProcessor(jon.Say);
            tomsVoice = new StringProcessor(tom.Say);
            background = new StringProcessor(Background.Note);

            jonsVoice("Hello, son.");
            tomsVoice.Invoke("Hello, Daddy!");
            background("An airplane flies past.");
        }
    }

}
