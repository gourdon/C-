using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Model;
using System.Xml.Linq;
using XmlExamples.Extensions;

namespace XmlExamples
{
    public class CreateXmlWithReflectionConcise
    {
        static void Main()
        {
            SampleData.AssignIDs();

            var projects = new XElement("projects",
                from project in SampleData.AllProjects
                select new XElement("project",
                    new { project.Name, project.ProjectID }
                    .AsXAttributes(),

                    from subscription in SampleData.AllSubscriptions
                    where subscription.Project == project
                    select new XElement("subscription",
                        new
                        {
                            subscription.EmailAddress,
                            subscription.NotificationSubscriptionID
                        }
                        .AsXAttributes())
                )
            );

            var users = new XElement("users",
                from user in SampleData.AllUsers
                select new XElement("user",
                    new { user.Name, user.UserID, user.UserType }
                    .AsXAttributes()
                )
            );

            var defects = new XElement("defects",
                from defect in SampleData.AllDefects
                select new XElement("defect",
                    new
                    {   defect.ID,
                        defect.Summary,
                        defect.Created,
                        AssignedTo = defect.AssignedTo == null ?
                                    null : (int?)defect.AssignedTo.UserID,
                        CreatedBy = defect.CreatedBy.UserID,
                        defect.Status,
                        defect.Severity,
                        defect.LastModified
                    }
                    .AsXAttributes()
                )
            );

            XElement doc = new XElement
            (
                new XElement("defect-system", projects, users, defects)
            );

            Console.WriteLine(doc.ToString());
        }
    }
}
