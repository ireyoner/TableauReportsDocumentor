using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace TableauReportsDocumentor.Modules.ImportModule
{
    class Import
    {
        public XmlDocument ImportTableauWorkbooks(string filename)
        {

            XmlDocument doc = new XmlDocument();
            XslCompiledTransform myXslTrans = new XslCompiledTransform();
            myXslTrans.Load("../../Import Converters/TRDCI_v8.3.xsl");

            XPathDocument myXPathDoc = new XPathDocument(filename);

            using (XmlWriter writer = doc.CreateNavigator().AppendChild())
            {
                myXslTrans.Transform(myXPathDoc, null, writer);
            }

            return doc;
        }

    }
}
