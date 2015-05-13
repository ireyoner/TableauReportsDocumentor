using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TableauReportsDocumentor.Export_Converters
{
    abstract class ExporterInstance
    {

        public bool Export(string exportFileName, System.Xml.XmlDocument exportSource)
        {
            XmlNode root = exportSource.DocumentElement;
            var title = root.SelectSingleNode("title").InnerText;
            
            ExportInit(title, root, exportFileName);

            foreach (XmlNode section in root.SelectNodes("sections/section"))
            {
                CreateSection(section);
            }

            return ExportEnd(title, root, exportFileName);
        }

        protected abstract void ExportInit(string reportTitle, XmlNode rootNode, string exportFileName);
        protected abstract bool ExportEnd(string reportTitle, XmlNode rootNode, string exportFileName);


        private void CreateSection(XmlNode section)
        {
            var title = section.SelectSingleNode("title").InnerText;

            SectionInit(title, section);

            SectionContentStart(title, section, section.SelectNodes("content/*").Count);
            foreach (XmlNode item in section.SelectNodes("content/*"))
            {
                if (item.Name == "text")
                {
                    CreateText(item);
                }
                else if (item.Name == "table")
                {
                    CreateTable(item);
                }
            }
            SectionContentEnd(title, section);

            SectionSubsectionsStart(title, section, section.SelectNodes("subsections/subsection").Count);
            foreach (XmlNode subsection in section.SelectNodes("subsections/subsection"))
            {
                CreateSubSection(subsection);
            }
            SectionSubsectionsEnd(title, section);

            SectionEnd(title, section);
        }

        protected abstract void SectionInit(string sectionTitle, XmlNode sectionNode);
        protected abstract void SectionContentStart(string sectionTitle, XmlNode sectionNode, int contentNodesCount);
        protected abstract void SectionContentEnd(string sectionTitle, XmlNode sectionNode);
        protected abstract void SectionSubsectionsStart(string sectionTitle, XmlNode sectionNode, int subsectionNodesCount);
        protected abstract void SectionSubsectionsEnd(string sectionTitle, XmlNode sectionNode);
        protected abstract void SectionEnd(string sectionTitle, XmlNode sectionNode);


        private void CreateSubSection(XmlNode subSection)
        {
            var title = subSection.SelectSingleNode("title").InnerText;

            SubSectionInit(title, subSection);

            SubSectionContentStart(title, subSection, subSection.SelectNodes("content/*").Count);
            foreach (XmlNode item in subSection.SelectNodes("content/*"))
            {
                if (item.Name == "text")
                {
                    CreateText(item);
                }
                else if (item.Name == "table")
                {
                    CreateTable(item);
                }
            }
            SubSectionContentEnd(title, subSection);

            SubSectionEnd(title, subSection);
        }

        protected abstract void SubSectionInit(string subSectionTitle, XmlNode subSectionNode);
        protected abstract void SubSectionContentStart(string subSectionTitle, XmlNode subSectionNode, int contentNodesCount);
        protected abstract void SubSectionContentEnd(string subSectionTitle, XmlNode subSectionNode);
        protected abstract void SubSectionEnd(string subSectionTitle, XmlNode subSectionNode);


        protected abstract void CreateText(XmlNode textNode);


        private void CreateTable(XmlNode table)
        {
            var title = table.SelectSingleNode("title").InnerText;

            TableInit(title, table);

            var header = table.SelectSingleNode("header");
            var headerCells = header.SelectNodes("hcell");
            int colCount = headerCells.Count;

            TableHeaderStart(title, header, colCount);
            foreach (XmlNode cell in header.SelectNodes("hcell"))
            {
                TableHeaderCell(title, cell);
            }
            TableHeaderEnd(title, header);

            var rows = table.SelectSingleNode("rows");
            var rowsRows = rows.SelectNodes("row");
            int rowsCount = rowsRows.Count;

            TableRowsStart(title, rows, rowsCount);
            foreach (XmlNode row in rowsRows)
            {
                var rowCells = row.SelectNodes("cell");
                int cellsCount = rowCells.Count;

                TableRowStart(title, row, cellsCount);
                foreach (XmlNode cell in rowCells)
                {
                    TableRowCell(title, cell);
                }
                TableRowEnd(title, row);
            }
            TableHeaderEnd(title, header);

            TableEnd(title, header);
        }

        protected abstract void TableInit(string tableTitle, XmlNode tableNode);
        protected abstract void TableHeaderStart(string tableTitle, XmlNode headerNode, int headerNodesCount);
        protected abstract void TableHeaderCell(string tableTitle, XmlNode cellNode);
        protected abstract void TableHeaderEnd(string tableTitle, XmlNode headerNode);
        protected abstract void TableRowsStart(string tableTitle, XmlNode rowsNode, int contentNodesCount);
        protected abstract void TableRowStart(string tableTitle, XmlNode rowNode, int contentNodesCount);
        protected abstract void TableRowCell(string tableTitle, XmlNode cellNode);
        protected abstract void TableRowEnd(string tableTitle, XmlNode rowNode);
        protected abstract void TableRowsEnd(string tableTitle, XmlNode rowsNode);
        protected abstract void TableEnd(string tableTitle, XmlNode tableNode);

    }
}
