using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace Chapter05
{
    [Description("Listing 5.12")]
    class CapturedVariableLifetimeExtension
    {
        static MethodInvoker CreateDelegateInstance()
        {
            int counter = 5;

            MethodInvoker ret = delegate
            {
                 Console.WriteLine(counter);
                 counter++;
            };
            ret();
            return ret;
        }

        static void Main()
        {
            MethodInvoker x = CreateDelegateInstance();
            x();
            x();
        }
    }
}
