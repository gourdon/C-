using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;
using System.Xml.Linq;

namespace Chapter14
{
    [Description("Listing 14.27 and 14.28")]
    class ExpandoXml
    {
        public static dynamic CreateDynamicXml(XElement element)
        {
            dynamic expando = new ExpandoObject();
            expando.XElement = element;
            expando.ToXml = (Func<string>)element.ToString;

            IDictionary<string, object> dictionary = expando;
            foreach (XElement subElement in element.Elements())
            {
                dynamic subNode = CreateDynamicXml(subElement);
                string name = subElement.Name.LocalName;
                string listName = name + "List";
                if (dictionary.ContainsKey(name))
                {
                    ((List<dynamic>) dictionary[listName]).Add(subNode);
                }
                else
                {
                    dictionary[name] = subNode;
                    dictionary[listName] = new List<dynamic> { subNode };
                }
            }
            return expando;
        }

        static void Main()
        {
            XDocument doc = XDocument.Load("books.xml");
            dynamic root = CreateDynamicXml(doc.Root);
            Console.WriteLine(root.book.author.ToXml());
            Console.WriteLine(root.bookList[2].excerpt.XElement.Value);
        }
    }
}
