using System;
using System.ComponentModel;
using System.Linq;
using System.Xml.Linq;

namespace XmlExamples
{
    [Description("Listing 12.9")]
    class ShowAllUsers
    {
        static void Main()
        {
            XElement root = XmlSampleData.GetElement();

            var query = root.Element("users").Elements()
                            .Select(user => new { Name=(string)user.Attribute("name"),
                                                  UserType=(string)user.Attribute("type") });

            foreach (var user in query)
            {
                Console.WriteLine ("{0}: {1}", user.Name, user.UserType);
            }
        }
    }
}
