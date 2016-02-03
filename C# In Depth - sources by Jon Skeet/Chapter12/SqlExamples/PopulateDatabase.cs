using System;
using System.ComponentModel;

using Model;

namespace SqlExamples
{
    class PopulateDatabase
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                context.Defects.InsertAllOnSubmit(SampleData.AllDefects);
                context.NotificationSubscriptions.InsertAllOnSubmit(SampleData.AllSubscriptions);
                context.Users.InsertAllOnSubmit(SampleData.AllUsers);
                context.Projects.InsertAllOnSubmit(SampleData.AllProjects);

                context.SubmitChanges();
            }
        }
    }
}
