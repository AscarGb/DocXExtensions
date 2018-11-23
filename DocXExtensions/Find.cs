using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;

namespace DocXExtensions
{
    public static class Find
    {
        public static bool FindFirstChild(this XContainer c, string elem, out XElement xElement)
        {
            var collection = c.Descendants(ComplexFields.ab + elem);

            var enumerator = collection.GetEnumerator();

            var r = enumerator.MoveNext();

            if (r)
            {
                xElement = enumerator.Current;
            }
            else
            {
                xElement = null;
            }

            return r;
        }
    }
}
