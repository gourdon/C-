using System;
using System.ComponentModel;
using System.Linq;
using System.Xml.Linq;
using Model;
using XmlExamples.Extensions;

namespace XmlExamples
{
    [Description("Listing 12.15")]
    class CreateUsersXmlWithExtension
    {
        static void Main()
        {
            var users = new XElement("users",
                from user in SampleData.AllUsers
                select new XElement("user",
                                    new { name = user.Name, type = user.UserType }
                                        .AsXAttributes()
                                   )
            );

            Console.WriteLine(users);
        }
    }
}
