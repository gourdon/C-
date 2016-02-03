using System;
using System.ComponentModel;
using System.Linq;

using Model;

namespace SqlExamples
{
    class DeferredLoading
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                Console.WriteLine("Loading defect");
                var defect = (from bug in context.Defects
                             where bug.Summary.Contains("Welsh")
                             select bug).Single();

                Console.WriteLine("Using Project property...");
                Console.WriteLine("Project={0}", defect.Project.Name);
            }
        }
    }
}
