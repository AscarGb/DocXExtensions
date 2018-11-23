using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xceed.Words.NET;

namespace DocXExtensions
{
    public class ComplexField
    {
        XNamespace ab = ComplexFields.ab;

        IEnumerable<XElement> _complexFieldSet;
        XElement _textElem;
        XElement _textElemParent;
        string _instrText;
        string _code;
        DocX _doc;

        public ComplexField(string code, string instrText, IEnumerable<XElement> complexFieldSet, XElement textElem, DocX doc)
        {
            _code = code;
            _complexFieldSet = complexFieldSet;
            _textElem = textElem;
            _textElemParent = _textElem.Parent;
            _instrText = instrText;
            _doc = doc;
        }
        public string RawCode
        {
            get { return _instrText; }            
        }
        public string Code
        {
            get
            {
                return _code;
            }            
        }
        /// <summary>
        /// Custom property value
        /// </summary>
        public string Value
        {
            get
            {
                if (_doc.CustomProperties.ContainsKey(Code))
                {
                    return _doc.CustomProperties[Code].Value.ToString();
                }
                else
                {
                    throw new Exception("Custom property not found");
                }
            }
            set
            {
                _doc.AddCustomProperty(new CustomProperty(Code, value));
            }
        }
        /// <summary>
        ///  Removes this complex field set from its parent.
        /// </summary>
        public void Remove()
        {
            foreach (var a in _complexFieldSet)
                a.Remove();
        }
        public void ReplaceWithValue()
        {
            ReplaceWith(Value);
        }
        public void ReplaceWith(string replacement)
        {
            var b_size = _textElemParent.FindFirstChild("b", out var e_b) ? e_b.Attribute(ab + "val").Value : "0";

            var szCs_size = _textElemParent.FindFirstChild("szCs", out var e_szCs) ? e_szCs.Attribute(ab + "val").Value : "28";

            var r = new XElement(XName.Get("r", ab.ToString()));
            var rPr = new XElement(XName.Get("rPr", ab.ToString()));

            var b = new XElement(XName.Get("b", ab.ToString()));
            b.SetAttributeValue(ab + "val", b_size);

            var szCs = new XElement(XName.Get("szCs", ab.ToString()));
            szCs.SetAttributeValue(ab + "val", szCs_size);

            var t = new XElement(XName.Get("t", ab.ToString())) { Value = replacement };

            r.Add(rPr);
            r.Add(t);
            rPr.Add(b);
            rPr.Add(szCs);

            _complexFieldSet.Last().AddAfterSelf(r);
            Remove();
        }
    }
}
