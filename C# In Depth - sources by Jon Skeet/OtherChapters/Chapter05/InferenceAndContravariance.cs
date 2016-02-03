using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace Chapter05
{
    [Description("Listing 5.02")]
    class InferenceAndContravariance
    {
        static void LogPlainEvent(object sender, EventArgs e)
        {
            Console.WriteLine ("An event occurred");
        }

        static void Main()
        {
            Button button = new Button();
            button.Text = "Click me";
            button.Click += LogPlainEvent;
            button.KeyPress += LogPlainEvent;
            button.MouseClick += LogPlainEvent;

            Form form = new Form();
            form.AutoSize = true;
            form.Controls.Add(button);
            Application.Run(form);
        }
    }
}
