using System;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Xml.Linq;
using System.Collections.Generic;

namespace XmlExamples.Extensions
{
    static class XmlExtensions
    {
        public static IEnumerable<XAttribute> AsXAttributes(this object data)
        {
            foreach (PropertyInfo prop in data.GetType().GetProperties())
            {
                object value = prop.GetValue(data, null);
                if (value != null)
                {
                    yield return new XAttribute(prop.Name.Replace("_", "-"), value);
                }
            }
        }
    }
}
