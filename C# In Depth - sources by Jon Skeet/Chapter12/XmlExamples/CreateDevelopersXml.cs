using System;
using System.ComponentModel;
using System.Linq;
using System.Xml.Linq;

using Model;

namespace XmlExamples
{
    [Description("Listing 12.8")]
    class CreateDevelopersXml
    {
        static void Main()
        {
            var developers = new XElement("developers",
                from user in SampleData.AllUsers
                where user.UserType == UserType.Developer
                select new XElement("developer", user.Name)
            );
            Console.WriteLine(developers);
        }
    }
}
