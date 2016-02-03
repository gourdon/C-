using System;
using System.Linq;

using Chapter11.Model;

namespace Chapter11.Queries
{
    class ShowAllClosedDefects
    {
        static void Main()
        {
            var query = from defect in SampleData.AllDefects
                        where defect.Status == Status.Closed
                        select defect;

            foreach (var defect in query)
            {
                Console.WriteLine(defect);
            }
        }
    }
}
