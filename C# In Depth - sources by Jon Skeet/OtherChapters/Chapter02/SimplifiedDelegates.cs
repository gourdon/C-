using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace Chapter02
{
    [Description("Listing 2.4")]
    class SimplifiedDelegates
    {
        static void HandleDemoEvent(object sender, EventArgs e)
        {
            Console.WriteLine("Handled by HandleDemoEvent");
        }

        static void Main()
        {
            EventHandler handler;

            handler = new EventHandler(HandleDemoEvent);
            handler(null, EventArgs.Empty);

            handler = HandleDemoEvent;
            handler(null, EventArgs.Empty);

            handler = delegate(object sender, EventArgs e)
                { Console.WriteLine("Handled anonymously"); };
            handler(null, EventArgs.Empty);

            handler = delegate
                { Console.WriteLine("Handled anonymously"); };
            handler(null, EventArgs.Empty);

            MouseEventHandler mouseHandler = HandleDemoEvent;
            mouseHandler(null, new MouseEventArgs(MouseButtons.None, 
                                                  0, 0, 0, 0));
        }
    }
}
