using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
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
                try
                {
                    convertedXml = CheckXml(value);
                }
                catch (Exception e3)
                {
                    convertedXml = value;
                    throw e3;
                }

            }
        }

        private XmlDocument CheckXml(XmlDocument xml)
        {
            if (xml == null)
                return xml;
            var validatedXml = ValidateReport(xml);
            if (validatedXml != null && "preprocessedReport".Equals(validatedXml.DocumentElement.LocalName))
            {
                //Console.WriteLine(GetXmlAsString(validatedXml));
                XmlNode replacements = validatedXml.SelectSingleNode("//replacements");
                XmlDocument reportxml = new XmlDocument();
                reportxml.LoadXml(validatedXml.SelectSingleNode("//report").OuterXml);

                String report = this.GetXmlAsString(reportxml);

                foreach (XmlNode replce in replacements.SelectNodes("replacement"))
                {
                    Boolean isRegexp = replce.Attributes["isRegexp"].InnerText.Equals("True");
                    String oldValue = replce.Attributes["original"].InnerText;
                    String newValue = replce.InnerText;

                    if (isRegexp)
                    {
                        report = Regex.Replace(report, oldValue, newValue);
                    }
                    else
                    {
                        report = report.Replace(oldValue, newValue);
                    }

                }
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(report);
                return ValidateReport(doc);
            }
            else
            {
                return validatedXml;
            }
        }

        private XmlDocument ValidateReport(XmlDocument xml)
        {
            xml.Schemas.Add("", Properties.Settings.Default.ReportDocumentValidator);
            xml.Validate(veh);
            return xml;
        }

        private String GetXmlAsString(XmlDocument xml)
        {
            if (xml != null)
            {
                MemoryStream mStream = new MemoryStream();
                XmlTextWriter writer = new XmlTextWriter(mStream, Encoding.Unicode);
                writer.Formatting = Formatting.Indented;
                writer.Indentation = 4;
                writer.QuoteChar = '\'';
                xml.WriteContentTo(writer);
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

        public String ConvertedXmlAsString
        {
            get
            {
                return GetXmlAsString(this.convertedXml);
            }
            set
            {
                if (value != null)
                {
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(value);
                    this.ConvertedXml = doc;
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
            try
            {
                this.ConvertedXml = converted;
            }
            catch (Exception e2)
            {
                this.convertedXml = converted;
                MessageBox.Show(e2.Message, "Open error!");
            }
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

            String delElemXpath = " //section[@visible='False']" +
                                  "|//subsection[@visible='False']" +
                                  "|//text[@visible='False']" +
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
