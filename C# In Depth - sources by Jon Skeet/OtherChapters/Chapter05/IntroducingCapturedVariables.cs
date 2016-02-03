using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace Chapter05
{
    [Description("Listing 5.10")]
    class IntroducingCapturedVariables
    {
        void EnclosingMethod()
        {
            int outerVariable = 5;                       
            string capturedVariable = "captured";        

            if (DateTime.Now.Hour==23)
            {
                int normalLocalVariable = DateTime.Now.Minute;       
                Console.WriteLine(normalLocalVariable);
            }

            MethodInvoker x = delegate()
            { 
                string anonLocal="local to anonymous method";    
                Console.WriteLine(capturedVariable + anonLocal); 
            };
            x();
        }

        static void Main()
        {
        }
    }
}
