using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Model;

namespace SqlExamples
{
    public class CleanDatabaseQuickly
    {
        static void Main()
        {
            using (var context = new DefectModelDataContext())
            {
                context.Log = Console.Out;

                // Delete the contents of the tables
                context.ExecuteCommand("DELETE FROM Defect");
                context.ExecuteCommand("DELETE FROM NotificationSubscription");
                context.ExecuteCommand("DELETE FROM Project");
                context.ExecuteCommand("DELETE FROM DefectUser");
                // Reset the identity columns
                context.ExecuteCommand("DBCC CHECKIDENT('Defect', RESEED, 0)");
                context.ExecuteCommand("DBCC CHECKIDENT('NotificationSubscription', RESEED, 0)");
                context.ExecuteCommand("DBCC CHECKIDENT('Project', RESEED, 0)");
                context.ExecuteCommand("DBCC CHECKIDENT('DefectUser', RESEED, 0)");
            }
        }
    }
}
