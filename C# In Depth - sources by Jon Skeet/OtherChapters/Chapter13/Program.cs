using System;
using MiscUtil;

namespace Chapter13
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
