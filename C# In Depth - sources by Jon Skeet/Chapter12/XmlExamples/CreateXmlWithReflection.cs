using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Model;
using System.Xml.Linq;
using XmlExamples.Extensions;

namespace XmlExamples
{
    public class CreateXmlWithReflection
    {
        static void Main()
        {
            SampleData.AssignIDs();

            var projects = new XElement("projects",
                from project in SampleData.AllProjects
                select new XElement ("project",
                    new { project=project.Name, 
                          id=project.ProjectID }
                    .AsXAttributes(),

                    from subscription in SampleData.AllSubscriptions
                    where subscription.Project == project
                    select new XElement("subscription",
                        new { email=subscription.EmailAddress,
                              id=subscription.NotificationSubscriptionID }
                        .AsXAttributes())
                )
            );

            var users = new XElement("users",
                from user in SampleData.AllUsers
                select new XElement("user",
                    new { name=user.Name,
                          id=user.UserID,
                          type=user.UserType }
                    .AsXAttributes()
                )
            );

            var defects = new XElement("defects",
                from defect in SampleData.AllDefects
                select new XElement("defect",
                    new { id = defect.ID,
                          summary=defect.Summary,
                          created=defect.Created,
                          assigned_to=defect.AssignedTo==null ? 
                                      null : (int?)defect.AssignedTo.UserID,
                          created_by=defect.CreatedBy.UserID,
                          status=defect.Status,
                          severity=defect.Severity,
                          last_modified=defect.LastModified }
                    .AsXAttributes()
                )
            );

            XElement doc = new XElement
            (
                new XElement("defect-system", projects, users, defects)
            );

            Console.WriteLine (doc.ToString());
        }
    }
}
