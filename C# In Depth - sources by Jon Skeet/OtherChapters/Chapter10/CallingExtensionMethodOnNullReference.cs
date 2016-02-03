using System;
using System.ComponentModel;

namespace Chapter10
{
    [Description("Listing 10.05")]
    class CallingExtensionMethodOnNullReference
    {
        static void Main()
        {
            object y = null;
            Console.WriteLine(y.IsNull());
            y = new object();
            Console.WriteLine(y.IsNull());
        }
    }
}
