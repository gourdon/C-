using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;

namespace Chapter05
{
    [Description("Listing 5.13")]
    class MultipleCaptures
    {
        static void Main()
        {
            List<MethodInvoker> list = new List<MethodInvoker>();

            for (int index = 0; index < 5; index++)
            {
                int counter = index * 10;
                list.Add(delegate
                {
                    Console.WriteLine(counter);
                    counter++;
                });
            }

            foreach (MethodInvoker t in list)
            {
                t();
            }

            list[0]();
            list[0]();
            list[0]();

            list[1]();
        }
    }
}
