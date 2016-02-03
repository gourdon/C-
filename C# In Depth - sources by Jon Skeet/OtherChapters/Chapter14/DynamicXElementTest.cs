using System;
using System.ComponentModel;
using System.Xml.Linq;

namespace Chapter14
{
    [Description("Listing 14.32")]
    class DynamicXElementTest
    {
        static void Main()
        {
            XDocument doc = XDocument.Load("books.xml");
            dynamic root = DynamicXElement.CreateInstance(doc.Root);

            Console.WriteLine(root.book[2]["name"]);
            Console.WriteLine(root.book[1].author[1]);
            Console.WriteLine(root.book);
        }
    }
}
