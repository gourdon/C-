using System;
using System.ComponentModel;
using WebForms = System.Web.UI.WebControls;
using WinForms = System.Windows.Forms;

namespace Chapter07
{
    [Description("Listing 7.05")]
    class SimpleAliases
    {        
        static void Main()
        {
            Console.WriteLine (typeof (WinForms.Button));
            Console.WriteLine (typeof (WebForms.Button));
        }
    }
}
