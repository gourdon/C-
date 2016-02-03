using System;
using System.ComponentModel;
using System.Xml.Linq;

namespace XmlExamples
{
    [Description("Listing 12.13")]
    class CreateSimpleXml
    {
        static void Main()
        {
            var root = new XElement("root",
                                    new XAttribute("attr", "value"),
                                    new XElement("child",
                                                 "text")
                                    );

            Console.WriteLine(root);
        }
    }
}
