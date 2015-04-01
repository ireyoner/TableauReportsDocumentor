using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;

namespace TableauReportsDocumentor.Export_Converters
{
    class DocxExport : ExportInterface
    {

        public string exportFormat
        {
            get { throw new NotImplementedException(); }
        }

        public bool export(string exportFileName, System.Xml.XmlDocument exportSource)
        {
            using (DocX document = DocX.Create(exportFileName))
            {
                Paragraph p = document.InsertParagraph("tests");

                document.Save();
            }
            //throw new NotImplementedException();
            return true;
        }
    }
}
