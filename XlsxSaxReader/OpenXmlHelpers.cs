using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace XlsxSaxReader
{
    internal class OpenXmlHelpers
    {
        internal static OpenXmlReader GetOpenXmlReader(WorksheetPart worksheetPart)
        {
            return OpenXmlReader.Create(worksheetPart);
        }

        internal static XlsxSheetDimensions GetDimensions(string path)
        {
            using (var spreadsheetDoc = SpreadsheetDocument.Open(path, false))
            {
                var sheet = spreadsheetDoc.WorkbookPart.Workbook.Descendants<Sheet>().First();
                var worksheetPart = (WorksheetPart)spreadsheetDoc.WorkbookPart.GetPartById(sheet.Id);

                using (var reader = GetOpenXmlReader(worksheetPart))
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
            }
        }

        internal static void SkipRows(OpenXmlReader reader, int page, int pageSize)
        {
            MoveReaderToFirstRow(reader);

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

        internal static List<List<string>> GetRows(int page, int pageSize, XlsxSheetDimensions dimensions, OpenXmlReader reader, Stylesheet styleSheet, SharedStringTable sharedStringTable)
        {
            if (reader.EOF) return new List<List<string>>();

            var result = new List<List<string>>(pageSize);
            do
            {
                var row = GetRow(page, pageSize, dimensions, reader, styleSheet, sharedStringTable);
                if (row == null)
                    break;

                result.Add(row);
            }
            while (reader.ReadNextSibling() && result.Count < pageSize);

            return result;
        }

        internal static void MoveReaderToFirstRow(OpenXmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.ElementType != typeof(Row)) continue;

                break;
            }
        }

        internal static List<string> GetRow(int page, int pageSize, XlsxSheetDimensions dimensions, OpenXmlReader reader, Stylesheet styleSheet, SharedStringTable sharedStringTable)
        {
            if (reader.EOF) return null;

            var rowValues = Enumerable.Repeat<string>(null, dimensions.MaxColNum).ToList();

            reader.ReadFirstChild();

            do
            {
                if (reader.ElementType != typeof(Cell)) continue;
                Cell c = (Cell)reader.LoadCurrentElement();

                string cellValue = GetCellValue(c, styleSheet, sharedStringTable);
                var colName = c.GetAttributes().First().Value;
                var index = GetColNum(colName) - 1;
                rowValues[index] = cellValue;
            } while (reader.ReadNextSibling());

            return rowValues;
        }

        internal static string GetCellValue(Cell excelCell, Stylesheet styleSheet, SharedStringTable sharedStringTable)
        {
            string value;
            if (excelCell == null ||
                string.IsNullOrWhiteSpace(excelCell.InnerText))
                return null;
            if (excelCell.DataType == null)
            {
                return GetCellValueWithoutConsideringDataType(excelCell, styleSheet);
            }

            value = excelCell.InnerText;
            //If none of the below cases are executed, return the innerText            
            switch (excelCell.DataType.Value)
            {
                case CellValues.String:
                    value = excelCell.CellValue.InnerText;
                    break;
                case CellValues.SharedString:
                    value = GetSharedStringItem(excelCell.CellValue, sharedStringTable);
                    break;
                case CellValues.Boolean:
                    switch (value)
                    {
                        case "0": value = "FALSE"; break;
                        default: value = "TRUE"; break;
                    }
                    break;
            }
            return value;
        }

        internal static string GetCellValueWithoutConsideringDataType(Cell excelCell, Stylesheet styleSheet)
        {
            CellFormat cellFormat = GetCellFormat(excelCell, styleSheet);
            if (cellFormat != null)
            {
                return GetFormatedValue(excelCell, cellFormat, styleSheet.NumberingFormats);
            }
            else
            {
                var num = double.Parse(excelCell.CellValue.InnerText, CultureInfo.InvariantCulture);
                return num.ToString(CultureInfo.InvariantCulture);
            }
        }

        internal static string GetFormatedValue(Cell cell, CellFormat cellformat, NumberingFormats numberingFormats)
        {
            string value = null;
            string format = null;

            if (!TryGetFormat(cellformat, numberingFormats, out format))
            {
                value = cell.InnerText;
            }
            else if (OpenXmlConstants.DateTimeNumberingFormats.Contains(cellformat.NumberFormatId))
            {
                var datetime = DateTime.FromOADate(double.Parse(cell.InnerText));
                value = datetime.ToString(format, CultureInfo.InvariantCulture);
            }
            else
            {
                var num = double.Parse(cell.InnerText, CultureInfo.InvariantCulture);
                value = num.ToString(format, CultureInfo.InvariantCulture);
            }

            return value;
        }

        internal static string GetSharedStringItem(CellValue cellValue, SharedStringTable sharedStringTable)
        {
            if (sharedStringTable == null ||
                sharedStringTable.Count == 0)
            {
                return null;
            }

            var index = int.Parse(cellValue.InnerText);
            var sharedStringItem = sharedStringTable.Elements<SharedStringItem>().ElementAt(index);
            return sharedStringItem.Text.Text;
        }

        internal static bool TryGetFormat(CellFormat cellformat, NumberingFormats numberingFormats, out string format)
        {
            format = null;
            return cellformat.NumberFormatId != 0 &&
                cellformat.ApplyNumberFormat != null &&
                cellformat.ApplyNumberFormat.Value &&
                (OpenXmlConstants.DefaultNumberingFormats.TryGetValue(cellformat.NumberFormatId, out format) ||
                TryGetNumberingFormatInStyles(cellformat.NumberFormatId, numberingFormats, out format));
        }

        internal static CellFormat GetCellFormat(Cell cell, Stylesheet styleSheet)
        {
            if (cell.StyleIndex == null)
                return null;

            int styleIndex = (int)cell.StyleIndex.Value;
            return styleSheet
                .CellFormats
                .ElementAt(styleIndex) as CellFormat;
        }

        internal static bool TryGetNumberingFormatInStyles(uint numberingFormatId, NumberingFormats numberingFormats, out string format)
        {
            format = null;

            if (numberingFormats == null ||
                numberingFormats.Count == 0)
            {
                return false;
            }

            var numberingFormat = numberingFormats
                .Elements<NumberingFormat>()
                .Where(i => i.NumberFormatId.Value == numberingFormatId)
                .FirstOrDefault();

            format = numberingFormat != null ? numberingFormat.FormatCode : null;
            return format != null;
        }

        internal static int GetColNum(string colName)
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

        internal static int GetRowCount(string endDimension)
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

        internal static int GetCharIndex(char c)
        {
            return c % 32;
        }
    }
}
