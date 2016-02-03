using System;
using System.IO;
using System.Windows.Forms;

using MiscUtil;

namespace SqlExamples
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            DirectoryInfo dir = new DirectoryInfo(Application.ExecutablePath);
            // Navigate all the way up to the top of the example code tree
            dir = dir.Parent.Parent.Parent.Parent.Parent;
            string databasesDirectory = Path.Combine(dir.FullName, "Databases");
            AppDomain.CurrentDomain.SetData("DataDirectory", databasesDirectory);

            ApplicationChooser.Run(typeof(Program), args);
        }
    }
}
