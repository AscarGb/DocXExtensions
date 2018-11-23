using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Xceed.Words.NET;

namespace DocXExtensions
{
    public static class ComplexFields
    {
        public static readonly XNamespace ab = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static IEnumerable<ComplexField> GetComplexFields(this Paragraph p, DocX doc)
        {
            IEnumerable<XElement> complexFields =
                    from el in p.Xml.Descendants(ab + "fldChar")
                    where (string)el.Attribute(ab + "fldCharType") == "begin"
                    select el;

            if (complexFields.Any())
            {

                foreach (var beginElement in complexFields.ToList().Where(a => a != null))
                {
                    XElement textElem = null;


                    string instrText = "";

                    var begin = beginElement.Parent;

                    List<XObject> complexFieldSet = new List<XObject> { begin };

                    var nexElement = begin.NextNode;

                    while (true)
                    {
                        if (nexElement is XElement)
                        {
                            var textElemt = (nexElement as XElement).Descendants(ab + "t");
                            if (textElemt.Any())
                            {
                                textElem = textElemt.First();
                            }

                            var instrs = (nexElement as XElement).Descendants(ab + "instrText");
                            if (instrs.Any())
                            {
                                //for multiline instrText in docx
                                instrText += instrs.First().Value.Trim();
                            }

                            var ends =
                            from el in (nexElement as XElement).Descendants(ab + "fldChar")
                            where (string)el.Attribute(ab + "fldCharType") == "end"
                            select el;

                            if (ends.Any())
                            {
                                complexFieldSet.Add(nexElement);
                                break;
                            }
                        }

                        complexFieldSet.Add(nexElement);
                        nexElement = nexElement.NextNode;
                    }


                    if (string.IsNullOrEmpty(instrText))
                    {
                        continue;
                    }


                    var arr = instrText.Trim()
                        .Replace("\\o", "")
                        .Replace("\\h", "")
                        .Replace("\\z", "")
                        .Replace("\\u", "")
                        .Replace("  ", " ").Split(' ');

                    if (arr.Length > 1)
                    {
                        var propname = arr[1];

                        if (doc.CustomProperties.ContainsKey(propname))
                        {
                            var prop_text = doc.CustomProperties[propname].Value.ToString();

                            var seq = complexFieldSet.Where(a => a is XElement).Select(a => (XElement)a);

                            yield return new ComplexField(propname, instrText, seq, textElem, doc);
                        }
                    }
                }
            }
        }
    }
}
