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
        public HashSet<XmlDocument> ImportTableauWorkbooks()
        {
            var documents = new HashSet<XmlDocument>();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Tableau Workbook (*.twb)|*.twb|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDialog.ShowDialog() == true)
            {
                XslCompiledTransform myXslTrans = new XslCompiledTransform();
                myXslTrans.Load("../../Import Converters/TRDCI_v8.3.xsl");
                foreach (string filename in openFileDialog.FileNames)
                {
                    XPathDocument myXPathDoc = new XPathDocument(filename);
                    XmlDocument doc = new XmlDocument();
                    using (XmlWriter writer = doc.CreateNavigator().AppendChild())
                    {
                        myXslTrans.Transform(myXPathDoc, null, writer);
                    }
                    documents.Add(doc);
                }
            }
            return documents;
        }

    }
}
