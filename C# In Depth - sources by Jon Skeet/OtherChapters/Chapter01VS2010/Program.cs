using System;
using System.IO;
using MiscUtil;

namespace Chapter01VS2010
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
