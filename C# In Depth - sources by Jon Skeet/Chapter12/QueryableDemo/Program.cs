using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using MiscUtil;

namespace QueryableDemo
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
