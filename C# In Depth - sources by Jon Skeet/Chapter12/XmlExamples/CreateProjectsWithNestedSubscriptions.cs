using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Model;
using System.ComponentModel;

namespace XmlExamples
{
    [Description("Listing 12.16")]
    class CreateProjectsWithNestedSubscriptions
    {
        static void Main()
        {
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

            Console.WriteLine(projects);
        }
    }
}
