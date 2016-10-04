using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSaxReader
{
    public class OpenXmlHelpers
    {
        public static OpenXmlReader GetOpenXmlReader(WorksheetPart worksheetPart)
        {
            return OpenXmlReader.Create(worksheetPart);
        }

        public static Dictionary<uint, string> GetCellFormats(SpreadsheetDocument spreadsheetDoc)
        {
            Dictionary<uint, string> formatMappings = new Dictionary<uint, string>();

            var stylePart = spreadsheetDoc.WorkbookPart.WorkbookStylesPart;

            var numFormatsParentNodes = stylePart.Stylesheet.ChildElements.OfType<NumberingFormats>();

            foreach (var numFormatParentNode in numFormatsParentNodes)
            {
                var formatNodes = numFormatParentNode.ChildElements.OfType<NumberingFormat>();
                foreach (var formatNode in formatNodes)
                {
                    formatMappings.Add(formatNode.NumberFormatId.Value, formatNode.FormatCode);
                }
            }

            return formatMappings;
        }

        public static XlsxSheetDimensions GetDimensions(OpenXmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.ElementType != typeof(SheetDimension)) continue;

                var sheetDimension = (SheetDimension)reader.LoadCurrentElement();
                var attr = sheetDimension.GetAttributes().First().Value;
                var dimensions = attr.Split(':');

                return new XlsxSheetDimensions(
                    GetRowCount(dimensions[0]),
                    GetRowCount(dimensions[1]),
                    GetColNum(dimensions[0]),
                    GetColNum(dimensions[1])
                    );
            }

            return null;
        }

        public static void SkipRows(OpenXmlReader reader, int page, int pageSize)
        {
            if (page == 0) return;
            if (pageSize == 0) return;

            var startIndex = (page - 1) * pageSize + 1;
            int rowNum = 0;
            do
            {
                if (reader.HasAttributes)
                {
                    rowNum = Convert.ToInt32(reader.Attributes.First(a => a.LocalName == "r").Value);
                }
            }
            while (Convert.ToInt32(rowNum) < startIndex && reader.ReadNextSibling());
        }

        public static bool TryGetFormat(SpreadsheetDocument spreadsheetDoc, CellFormat cellformat, out string format)
        {
            format = null;
            return cellformat.NumberFormatId != 0 &&
                cellformat.ApplyNumberFormat != null &&
                cellformat.ApplyNumberFormat.Value &&
                (OpenXmlConstants.DefaultNumberingFormats.TryGetValue(cellformat.NumberFormatId, out format) ||
                TryGetNumberingFormatInStyles(spreadsheetDoc, cellformat.NumberFormatId, out format));
        }

        #region private helpers

        private static CellFormat GetCellFormat(SpreadsheetDocument spreadsheetDoc, Cell cell)
        {
            if (cell.StyleIndex == null ||
                !HasCellFormats(spreadsheetDoc))
                return null;

            int styleIndex = (int)cell.StyleIndex.Value;
            return (CellFormat)spreadsheetDoc
                .WorkbookPart
                .WorkbookStylesPart
                .Stylesheet
                .CellFormats
                .ElementAt(styleIndex);
        }

        private static bool TryGetNumberingFormatInStyles(SpreadsheetDocument spreadsheetDoc, uint numberingFormatId, out string format)
        {
            format = null;

            if (!HasNumberingFormats(spreadsheetDoc)) return false;

            var numberingFormat = spreadsheetDoc
                .WorkbookPart
                .WorkbookStylesPart
                .Stylesheet
                .NumberingFormats
                .Elements<NumberingFormat>()
                .Where(i => i.NumberFormatId.Value == numberingFormatId)
                .FirstOrDefault();

            format = numberingFormat != null ? numberingFormat.FormatCode : null;
            return format != null;
        }

        private static int GetColNum(string colName)
        {
            var colNum = 1;
            foreach (var c in colName)
            {
                if (!char.IsLetter(c))
                    break;

                colNum *= GetCharIndex(c);
            }
            return colNum;
        }

        private static int GetRowCount(string endDimension)
        {
            var rowCount = 0;
            for (int i = 0; i < endDimension.Length; i++)
            {
                if (char.IsLetter(endDimension[i]))
                    continue;

                rowCount = Convert.ToInt32(endDimension.Substring(i, endDimension.Length - i));
                break;
            }
            return rowCount;
        }

        private static int GetCharIndex(char c)
        {
            return c % 32;
        }

        private static bool HasCellFormats(SpreadsheetDocument spreadsheetDoc)
        {
            return HasStylesheet(spreadsheetDoc) &&
                spreadsheetDoc.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats != null &&
                spreadsheetDoc.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count > 0;
        }

        private static bool HasNumberingFormats(SpreadsheetDocument spreadsheetDoc)
        {
            return HasStylesheet(spreadsheetDoc) &&
                spreadsheetDoc.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats != null &&
                spreadsheetDoc.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Count > 0;
        }

        private static bool HasStylesheet(SpreadsheetDocument spreadsheetDoc)
        {
            return HasWorkbookStylesPart(spreadsheetDoc) &&
                spreadsheetDoc.WorkbookPart.WorkbookStylesPart.Stylesheet != null;
        }

        private static bool HasWorkbookStylesPart(SpreadsheetDocument spreadsheetDoc)
        {
            return HasWorkbookPart(spreadsheetDoc) &&
                spreadsheetDoc.WorkbookPart.WorkbookStylesPart != null;
        }

        private static bool HasWorkbookPart(SpreadsheetDocument spreadsheetDoc)
        {
            return spreadsheetDoc.WorkbookPart != null;
        }

        #endregion
    }
}
