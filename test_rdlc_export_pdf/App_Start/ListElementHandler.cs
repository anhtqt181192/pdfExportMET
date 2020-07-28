using iTextSharp.text;
using iTextSharp.tool.xml;
using iTextSharp.tool.xml.pipeline;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace test_rdlc_export_pdf.App_Start
{
    public class ListElementHandler : IElementHandler
    {
        List<IElement> elements = new List<IElement>();

        public List<IElement> List => elements;

        public void Add(IWritable w)
        {
            if (w is WritableElement)
            {
                foreach (IElement e in ((WritableElement)w).Elements())
                {
                    elements.Add(e);
                }
            }
        }
    }
}