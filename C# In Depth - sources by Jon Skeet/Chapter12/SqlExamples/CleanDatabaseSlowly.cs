using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Model;

namespace SqlExamples
{
    public class CleanDatabaseSlowly
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                context.NotificationSubscriptions.DeleteAllOnSubmit(from subscription in context.NotificationSubscriptions select subscription);
                context.Projects.DeleteAllOnSubmit(from project in context.Projects select project);
                context.Defects.DeleteAllOnSubmit(from defect in context.Defects select defect);
                context.Users.DeleteAllOnSubmit(from user in context.Users select user);

                context.SubmitChanges();
            }
        }
    }
}
