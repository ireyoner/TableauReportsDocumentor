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
        public XmlDocument ImportTableauWorkbooks()
        {

            XmlDocument doc = new XmlDocument();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Tableau Workbook (*.twb)|*.twb|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDialog.ShowDialog() == true)
            {
                XslCompiledTransform myXslTrans = new XslCompiledTransform();
                myXslTrans.Load("../../Import Converters/TRDCI_v8.3.xsl");
                string filename = openFileDialog.FileName;
                
                    XPathDocument myXPathDoc = new XPathDocument(filename);
                    
                    using (XmlWriter writer = doc.CreateNavigator().AppendChild())
                    {
                        myXslTrans.Transform(myXPathDoc, null, writer);
                    }
     
                
            }
            return doc;
        }

    }
}
