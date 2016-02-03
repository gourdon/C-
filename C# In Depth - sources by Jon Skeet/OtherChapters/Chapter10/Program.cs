using System;
using MiscUtil;

namespace Chapter10
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
