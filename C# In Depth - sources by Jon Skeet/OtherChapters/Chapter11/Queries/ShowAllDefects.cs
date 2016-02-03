using System;
using System.Linq;

using Chapter11.Model;

namespace Chapter11.Queries
{
    class ShowAllDefects
    {
        static void Main()
        {
            var query = from defect in SampleData.AllDefects
                        select defect;

            foreach (var defect in query)
            {
                Console.WriteLine(defect);
            }
        }
    }
}
