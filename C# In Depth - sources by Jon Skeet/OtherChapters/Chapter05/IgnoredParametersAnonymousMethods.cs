using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace Chapter05
{
    [Description("Listing 5.09")]
    class IgnoredParametersAnonymousMethods
    {
        static void Main()
        {
            Button button = new Button();
            button.Text = "Click me";
            button.Click += delegate { Console.WriteLine("LogPlain"); };
            button.KeyPress += delegate { Console.WriteLine("LogKey"); };
            button.MouseClick += delegate { Console.WriteLine("LogMouse"); };

            Form form = new Form();
            form.AutoSize = true;
            form.Controls.Add(button);
            Application.Run(form);
        }
    }
}
