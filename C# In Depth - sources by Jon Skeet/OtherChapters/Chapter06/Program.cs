using System;
using MiscUtil;

namespace Chapter06
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
