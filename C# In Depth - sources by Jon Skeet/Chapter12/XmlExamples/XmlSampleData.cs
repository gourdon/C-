using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Model;

namespace XmlExamples
{
    public static class XmlSampleData
    {
        public static XElement GetElement()
        {
            SampleData.AssignIDs();

            var projects = new XElement("projects",
                from project in SampleData.AllProjects
                select new XElement("project",
                    new XAttribute("name", project.Name),
                    new XAttribute("id", project.ProjectID),
                    from subscription in SampleData.AllSubscriptions
                    where subscription.Project == project
                    select new XElement("subscription",
                        new XAttribute("email", subscription.EmailAddress)
                    )
                )
            );

            var users = new XElement("users",
                from user in SampleData.AllUsers
                select new XElement("user",
                    new XAttribute("name", user.Name),
                    new XAttribute("id", user.UserID),
                    new XAttribute("type", user.UserType))
            );

            var defects = new XElement("defects",
                from defect in SampleData.AllDefects
                select new XElement("defect",
                    new XAttribute("id", defect.ID),
                    new XAttribute("summary", defect.Summary),
                    new XAttribute("created", defect.Created),
                    new XAttribute("project", defect.Project.ProjectID),
                    defect.AssignedTo == null ? null : new XAttribute("assigned-to", defect.AssignedTo.UserID),
                    new XAttribute("created-by", defect.CreatedByUserID),
                    new XAttribute("status", defect.Status),
                    new XAttribute("severity", defect.Severity),
                    new XAttribute("last-modified", defect.LastModified))
            );

            XElement root = new XElement("defect-system", projects, users, defects);

            return root;
        }
    }
}
