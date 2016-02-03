using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace Chapter05
{
    [Description("Listing 5.11")]
    class CapturedVariableWalkthrough
    {
        static void Main()
        {
            string captured = "before x is created";

            MethodInvoker x = delegate
            {
                Console.WriteLine(captured);
                captured = "changed by x";
            };

            captured = "directly before x is invoked";
            x();

            Console.WriteLine(captured);

            captured = "before second invocation";
            x();
        }
    }
}
