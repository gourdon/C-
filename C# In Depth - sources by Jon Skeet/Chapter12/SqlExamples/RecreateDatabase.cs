using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Model;

namespace SqlExamples
{
    public class RecreateDatabase
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                context.DeleteDatabase();
                context.CreateDatabase();
            }
        }
    }
}
