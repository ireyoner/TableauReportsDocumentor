using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;

namespace TableauReportsDocumentor.Data
{
    public class ReportContent
    {

        private XmlDocument convertedXml;
        public XmlDocument ConvertedXml
        {
            get
            {
                return convertedXml;
            }
            set
            {
                if (value != null)
                {
                    convertedXml = value;
                    try
                    {
                        convertedXml.Schemas.Add("", "../../Import Converters/ImportValidator.xsd");
                        convertedXml.Validate(veh);
                    }
                    catch (Exception e)
                    {
                        throw e;
                    }
                }
            }
        }
        public String ConvertedXmlAsString
        {
            get
            {
                if (this.convertedXml != null)
                {
                    MemoryStream mStream = new MemoryStream();
                    XmlTextWriter writer = new XmlTextWriter(mStream, Encoding.Unicode);
                    writer.Formatting = Formatting.Indented;
                    writer.Indentation = 4;
                    writer.QuoteChar = '\'';
                    this.convertedXml.WriteContentTo(writer);
                    writer.Flush();
                    mStream.Flush();
                    mStream.Position = 0;
                    StreamReader sReader = new StreamReader(mStream);
                    String FormattedXML = sReader.ReadToEnd();
                    return FormattedXML;
                }
                else
                {
                    return "";
                }
            }
            set
            {
                if (value != null)
                {
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(value);
                    this.convertedXml = doc;
                }
            }
        }

        public String Original { get; private set; }

        public ReportContent()
        {
            SetupReportContent(null, null);
        }

        public ReportContent(String original)
        {
            SetupReportContent(original, null);
        }

        public ReportContent(XmlDocument converted)
        {
            SetupReportContent(null, converted);
        }

        public ReportContent(String original, XmlDocument converted)
        {
            SetupReportContent(original, converted);
        }

        private void SetupReportContent(String original, XmlDocument converted)
        {
            this.ConvertedXml = converted;
            this.Original = original;
        }

        private ValidationEventHandler veh;
        private static void ValidationCallBack(object sender, ValidationEventArgs args)
        {
            String message = "";
            if (args.Severity == XmlSeverityType.Warning)
            {
                message = "\tWarning: Matching schema not found.  No validation occurred." + args.Message;
            }
            else
            {
                message = "\tValidation error: " + args.Message;
            }
            Console.WriteLine(message);
            throw new Exception("Document validation exception!\n" + message);
        }

        public XmlDocument GetExportXml()
        {
            XmlDocument doc = (XmlDocument)this.ConvertedXml.Clone();

            String delElemXpath = " //section[@visible='False']"+
                                  "|//subsection[@visible='False']"+
                                  "|//text[@visible='False']"+
                                  "|//table[@visible='False']";
            var elem = doc.SelectNodes(delElemXpath);
            while (elem.Count > 0)
            {
                elem[0].ParentNode.RemoveChild(elem[0]);
                elem = doc.SelectNodes(delElemXpath);
            }

            String delColXpath = "//table[header/hcell/@visible='False']";
            XmlNodeList tableList = doc.SelectNodes(delColXpath);
            while (tableList.Count > 0)
            {
                XmlNode table = tableList[0];
                XmlNode tableHeader = tableList[0].SelectSingleNode("header");
                XmlNodeList allHcells = table.SelectNodes("header/hcell");
                for (int i = allHcells.Count - 1; i >= 0; i--)
                {
                    if (allHcells[i].Attributes["visible"].Value == "False")
                    {
                        tableHeader.RemoveChild(allHcells[i]);
                        foreach (XmlNode row in table.SelectNodes("rows/row"))
                        {
                            row.RemoveChild(row.SelectNodes("cell")[i]);
                        }
                    }
                }

                tableList = doc.SelectNodes(delColXpath);
            }

            return doc;
        }
    }
}
