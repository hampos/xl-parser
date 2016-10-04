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

        public static void MoveToFirstRow(OpenXmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.ElementType != typeof(Row)) continue;

                break;
            }
        }

        public static Dictionary<uint, string> GetCellFormats(string path)
        {
            using (var spreadsheetDoc = SpreadsheetDocument.Open(path, false))
            {
                return GetCellFormats(spreadsheetDoc.WorkbookPart);
            }
        }

        public static Dictionary<uint, string> GetCellFormats(WorkbookPart workbookpart)
        {
            Dictionary<uint, string> formatMappings = new Dictionary<uint, string>();

            var stylePart = workbookpart.WorkbookStylesPart;

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

        #region private helpers

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

        #endregion
    }
}
