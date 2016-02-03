using System;
using System.Linq;
using Chapter11.Model;

namespace Chapter10
{
    class DefectReport
    {
        static void Main()
        {
            var bugs = SampleData.AllDefects.Where(defect => defect.AssignedTo != null &&
                                                             defect.AssignedTo.UserType==UserType.Developer);
            
            var query = bugs.GroupBy(bug => bug.AssignedTo)
                            .Select(list => new { Developer = list.Key, Count = list.Count() })
                            .OrderByDescending(x => x.Count);

            foreach (var item in query)
            {
                Console.WriteLine(item);
            }
        }
    }
}
