using System;
using System.ComponentModel;
using System.Linq;
using System.Xml.Linq;

using Model;

namespace XmlExamples
{
    [Description("Listing 12.7")]
    class CreateUsersXml
    {
        static void Main()
        {
            var users = new XElement("users", 
                SampleData.AllUsers.Select(user => new XElement("user",
                        new XAttribute("name", user.Name),
                        new XAttribute("type", user.UserType)))
            );

            Console.WriteLine(users);
        }
    }
}
