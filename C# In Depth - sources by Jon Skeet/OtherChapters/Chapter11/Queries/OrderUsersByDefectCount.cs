using System;
using System.Linq;

using Chapter11.Model;

namespace Chapter11.Queries
{
    class OrderUsersByDefectCount
    {
        static void Main()
        {
            var query = from user in SampleData.AllUsers
                        join defect in SampleData.AllDefects on user equals defect.AssignedTo into defects
                        let count = defects.Count()
                        orderby count descending, user.Name
                        select new { User = user, Count = count };

            foreach (var userAndCount in query)
            {
                Console.WriteLine("{0}: {1}", userAndCount.User.Name, userAndCount.Count);
            }
        }
    }
}
