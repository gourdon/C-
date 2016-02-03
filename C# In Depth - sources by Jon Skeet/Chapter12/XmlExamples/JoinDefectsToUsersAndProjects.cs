using System;
using System.ComponentModel;
using System.Linq;
using System.Xml.Linq;

namespace XmlExamples
{
    [Description("Listing 12.18")]
    class JoinDefectsToUsersAndProjects
    {
        static void Main()
        {
            XElement root = XmlSampleData.GetElement();

            var query = from defect in root.Descendants("defect")
                        join user in root.Descendants("user")
                          on (int?)defect.Attribute("assigned-to") equals
                             (int)user.Attribute("id")
                        join project in root.Descendants("project")
                          on (int)defect.Attribute("project") equals
                             (int)project.Attribute("id")
                        where (string)defect.Attribute("status") != "Closed"
                        select new { ID=(int)defect.Attribute("id"),
                                     Project=(string)project.Attribute("name"),
                                     Assignee=(string)user.Attribute("name") };

            foreach (var defect in query)
            {
                Console.WriteLine ("{0}: {1}/{2}", 
                                   defect.ID, 
                                   defect.Project, 
                                   defect.Assignee);
            }                             
        }
    }
}
