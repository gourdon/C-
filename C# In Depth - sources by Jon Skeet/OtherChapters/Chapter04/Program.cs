using System;
using MiscUtil;

namespace Chapter04
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            ApplicationChooser.Run(typeof(Program), args);
        }
    }
}
