using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Model;
using System.ComponentModel;

namespace SqlExamples
{
    class UpdateFirstDefect
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                Defect defect = context.Defects
                                       .Where(d => d.ID==1)
                                       .Single();

                User tim = defect.CreatedBy;

                defect.AssignedTo = tim;
                tim.Name = "Timothy Trotter";
                defect.Status = Status.Fixed;
                defect.LastModified = SampleData.May(31);

                context.SubmitChanges();
            }

            using (var context = new DefectModelDataContext())
            {
                Defect d = (from defect in context.Defects
                            where defect.ID==1
                            select defect).Single();

                Console.WriteLine (d);
            }
        }
    }
}
