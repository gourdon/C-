using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Model;

namespace SqlExamples
{
    class ShowAllClosedDefects
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                var query = from defect in context.Defects
                            where defect.Status == Status.Closed
                            select defect;

                foreach (var defect in query)
                {
                    Console.WriteLine(defect);
                }
            }
        }
    }
}
