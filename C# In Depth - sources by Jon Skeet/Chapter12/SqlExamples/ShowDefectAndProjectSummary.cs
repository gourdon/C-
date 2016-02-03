using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Model;

namespace SqlExamples
{
    class ShowDefectAndProjectSummary
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                var query = from defect in context.Defects
                            select new { defect.Summary, ProjectName=defect.Project.Name };

                foreach (var entry in query)
                {
                    Console.WriteLine("{0}: {1}", entry.ProjectName, entry.Summary);
                }
            }
        }
    }
}
