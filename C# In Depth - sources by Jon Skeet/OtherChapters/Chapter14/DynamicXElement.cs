using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;
using System.Linq;
using System.Xml.Linq;

namespace Chapter14
{
    [Description("Listing 14.29/30/31/33")]
    public sealed class DynamicXElement : DynamicObject
    {
        private readonly XElement element;

        private DynamicXElement(XElement element)
        {
            this.element = element;
        }

        public static dynamic CreateInstance(XElement element)
        {
            return new DynamicXElement(element);
        }

        public override string ToString()
        {
            return element.ToString();
        }

        public XElement XElement
        {
            get { return element; }
        }

        public XAttribute this[XName name]
        {
            get { return element.Attribute(name); }
        }

        public dynamic this[int index]
        {
            get
            {
                XElement parent = element.Parent;
                if (parent == null && index != 0)
                {
                    throw new ArgumentOutOfRangeException();
                }
                return new DynamicXElement(parent.Elements(element.Name).ElementAt(index));
            }
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            string name = binder.Name;
            XElement subElement = element.Element(name);
            if (subElement != null)
            {
                result = new DynamicXElement(subElement);
                return true;
            }
            return base.TryGetMember(binder, out result);
        }

        public override IEnumerable<string> GetDynamicMemberNames()
        {
            return element.Elements()
                          .Select(x => x.Name.LocalName)
                          .Distinct()
                          .OrderBy(x => x);
        }
    }
}
