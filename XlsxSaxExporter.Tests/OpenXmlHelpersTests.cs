using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Xunit;

namespace XlsxSaxExporter.Tests
{
    public class OpenXmlHelpersTests
    {
        [Fact]
        public void Given_WorksheetPart_I_Can_Create_OpenXmlReader()
        {
            var temp = Path.GetTempFileName();
            TestHelpers.Create(temp);

            using (var spreadsheetDocument = SpreadsheetDocument.Open(temp, false))
            {
                var worksheetPart = spreadsheetDocument.WorkbookPart.WorksheetParts.First();

                var reader = OpenXmlHelpers.GetOpenXmlReader(worksheetPart);

                reader.Should().NotBeNull();
            }

            File.Delete(temp);
        }

        [Fact]
        public void Given_Xlsx_Path_I_Can_Get_Sheet_Dimensions()
        {
            var temp = Path.GetTempFileName();
            TestHelpers.Create(temp, sheetDimensionRef: "A1:B2");

            var dimensions = OpenXmlHelpers.GetDimensions(temp);

            dimensions.Should().NotBeNull();
            dimensions.MinRowNum.Should().Be(1);
            dimensions.MaxRowNum.Should().Be(2);
            dimensions.MinColNum.Should().Be(1);
            dimensions.MaxColNum.Should().Be(2);

            File.Delete(temp);
        }

        [Fact]
        public void Given_Cell_Reference_I_Can_Get_Col_Num()
        {
            string dimensions = "B1";

            var rowCount = OpenXmlHelpers.GetColNum(dimensions);

            rowCount.Should().Be(2);
        }

        [Fact]
        public void Given_Cell_Reference_I_Can_Get_Row_Num()
        {
            string dimensions = "A2";

            var rowCount = OpenXmlHelpers.GetRowNum(dimensions);

            rowCount.Should().Be(2);
        }

        [Fact]
        public void Given_NumberingFormatId_And_NumberingFormats_When_It_Exists_In_NumberingFormats_Then_TryGetNumberingFormatInStyles_Sets_Format_And_Returns_True()
        {
            var numberingFormat = new NumberingFormat
            {
                NumberFormatId = 1,
                FormatCode = "test"
            };
            var numberingFormats = new NumberingFormats();
            numberingFormats.AppendChild(numberingFormat);
            numberingFormats.Count = 1;

            string format = null;
            var success = OpenXmlHelpers.TryGetNumberingFormatInStyles(1, numberingFormats, out format);

            success.Should().BeTrue();
            format.Should().Be("test");
        }

        [Fact]
        public void Given_NumberingFormatId_And_NumberingFormats_When_It_Doesnt_Exist_In_NumberingFormats_Then_TryGetNumberingFormatInStyles_Returns_False()
        {
            var numberingFormat = new NumberingFormat
            {
                NumberFormatId = 1,
                FormatCode = "test"
            };
            var numberingFormats = new NumberingFormats();
            numberingFormats.AppendChild(numberingFormat);
            numberingFormats.Count = 1;

            string format = null;
            var success = OpenXmlHelpers.TryGetNumberingFormatInStyles(2, numberingFormats, out format);

            success.Should().BeFalse();
            format.Should().BeNull();
        }

        [Fact]
        public void Given_NumberingFormatId_And_NumberingFormats_When_NumberingFormats_Is_Null_Then_TryGetNumberingFormatInStyles_Returns_False()
        {
            string format = null;
            var success = OpenXmlHelpers.TryGetNumberingFormatInStyles(1, null, out format);

            success.Should().BeFalse();
            format.Should().BeNull();
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_And_Cell_StyleIndex_Is_Null_Then_GetFormat_Returns_Null()
        {
            var cell = new Cell();
            var stylesheet = new Stylesheet();

            var result = OpenXmlHelpers.GetCellFormat(cell, stylesheet);

            result.Should().BeNull();
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_And_CellFormat_Exists_For_Cell_Then_GetFormat_Returns_CellFormat()
        {
            var cellFormat = new CellFormat
            {
                FormatId = 0
            };
            var cell = new Cell()
            {
                StyleIndex = cellFormat.FormatId
            };
            var cellFormats = new CellFormats(cellFormat);
            var stylesheet = new Stylesheet(cellFormats);

            var result = OpenXmlHelpers.GetCellFormat(cell, stylesheet);

            result.ShouldBeEquivalentTo(cellFormat);
        }

        [Fact]
        public void Given_CellFormat_With_NumberingFormat_Zero_Then_TryGetFormat_Returns_False()
        {
            var cellFormat = new CellFormat
            {
                NumberFormatId = 0
            };
            var numberingFormats = new NumberingFormats();

            string format = null;

            var success = OpenXmlHelpers.TryGetFormat(cellFormat, numberingFormats, out format);

            success.Should().BeFalse();
            format.Should().BeNull();
        }

        [Fact]
        public void Given_CellFormat_With_ApplyNumberFormat_Null_Then_TryGetFormat_Returns_False()
        {
            var cellFormat = new CellFormat
            {
                NumberFormatId = 1
            };
            var numberingFormats = new NumberingFormats();

            string format = null;

            var success = OpenXmlHelpers.TryGetFormat(cellFormat, numberingFormats, out format);

            success.Should().BeFalse();
            format.Should().BeNull();
        }

        [Fact]
        public void Given_CellFormat_With_ApplyNumberFormat_False_Then_TryGetFormat_Returns_False()
        {
            var cellFormat = new CellFormat
            {
                NumberFormatId = 1,
                ApplyNumberFormat = false
            };
            var numberingFormats = new NumberingFormats();

            string format = null;

            var success = OpenXmlHelpers.TryGetFormat(cellFormat, numberingFormats, out format);

            success.Should().BeFalse();
            format.Should().BeNull();
        }

        [Fact]
        public void Given_CellFormat_With_Non_Default_NumberingFormat_And_NumberingFormat_Not_In_NumberingFormats_Then_TryGetFormat_Returns_False()
        {
            var cellFormat = new CellFormat
            {
                NumberFormatId = 164,
                ApplyNumberFormat = true
            };
            var numberingFormats = new NumberingFormats();

            string format = null;

            var success = OpenXmlHelpers.TryGetFormat(cellFormat, numberingFormats, out format);

            success.Should().BeFalse();
            format.Should().BeNull();
        }

        [Fact]
        public void Given_CellFormat_With_Default_NumberingFormat_Then_TryGetFormat_Returns_True_And_Sets_Format()
        {
            var cellFormat = new CellFormat
            {
                NumberFormatId = 1,
                ApplyNumberFormat = true
            };
            var numberingFormats = new NumberingFormats();

            string format = null;

            var success = OpenXmlHelpers.TryGetFormat(cellFormat, numberingFormats, out format);

            success.Should().BeTrue();
            format.Should().Be(OpenXmlConstants.DefaultNumberingFormats[1]);
        }

        [Fact]
        public void Given_CellFormat_With_Non_Default_NumberingFormat_And_NumberingFormat_In_NumberingFormats_Then_TryGetFormat_Returns_True_And_Sets_Format()
        {
            var cellFormat = new CellFormat
            {
                NumberFormatId = 164,
                ApplyNumberFormat = true
            };

            var numberingFormat = new NumberingFormat
            {
                NumberFormatId = 164,
                FormatCode = "test"
            };
            var numberingFormats = new NumberingFormats();
            numberingFormats.AppendChild(numberingFormat);
            numberingFormats.Count = 1;

            string format = null;

            var success = OpenXmlHelpers.TryGetFormat(cellFormat, numberingFormats, out format);

            success.Should().BeTrue();
            format.Should().Be("test");
        }

        [Fact]
        public void Given_Cell_And_CellFormat_And_NumberingFormats_When_Cannot_Get_Format_Returns_Cell_InnerText()
        {
            var cell = new Cell
            {
                CellValue = new CellValue("10")
            };
            var cellFormat = new CellFormat();
            var numberingFormats = new NumberingFormats();

            var result = OpenXmlHelpers.GetFormatedValue(cell, cellFormat, numberingFormats);

            result.Should().Be(cell.InnerText);
        }

        [Fact]
        public void Given_Cell_And_CellFormat_And_NumberingFormats_When_CellFormat_Exists_And_NumberingFormat_Is_Default_Then_GetFormatedValue_Returns_Formatted_Cell_InnerText()
        {
            var cell = new Cell
            {
                CellValue = new CellValue("10,00")
            };
            var cellFormat = new CellFormat
            {
                FormatId = 0,
                ApplyNumberFormat = true,
                NumberFormatId = 1
            };
            var numberingFormats = new NumberingFormats();
            var expectedValue = double.Parse(cell.CellValue.InnerText, CultureInfo.InvariantCulture).ToString(OpenXmlConstants.DefaultNumberingFormats[1]); 

            var result = OpenXmlHelpers.GetFormatedValue(cell, cellFormat, numberingFormats);

            result.Should().Be(expectedValue);
        }

        [Fact]
        public void Given_Cell_And_CellFormat_And_NumberingFormats_When_CellFormat_Is_Date_Then_GetFormatedValue_Returns_Formatted_Cell_InnerText()
        {
            var date = DateTime.UtcNow;
            var dateDoubleValue = date.ToOADate();
            var dateNumberingFormatId = OpenXmlConstants.DateTimeNumberingFormats[0];
            var expectedValue = date.ToString(OpenXmlConstants.DefaultNumberingFormats[dateNumberingFormatId], CultureInfo.InvariantCulture);

            var cell = new Cell
            {
                CellValue = new CellValue(dateDoubleValue.ToString())
            };
            var cellFormat = new CellFormat
            {
                FormatId = 0,
                ApplyNumberFormat = true,
                NumberFormatId = dateNumberingFormatId
            };
            var numberingFormats = new NumberingFormats();

            var result = OpenXmlHelpers.GetFormatedValue(cell, cellFormat, numberingFormats);

            result.Should().Be(expectedValue);
        }

        [Fact]
        public void Given_Cell_And_CellFormat_And_NumberingFormats_When_CellFormat_Exists_And_NumberingFormat_In_NumberingFormats_Then_GetFormatedValue_Returns_Formatted_Cell_InnerText()
        {
            var cell = new Cell
            {
                CellValue = new CellValue("10")
            };
            var cellFormat = new CellFormat
            {
                FormatId = 0,
                ApplyNumberFormat = true,
                NumberFormatId = 164
            };
            var numberingFormat = new NumberingFormat
            {
                NumberFormatId = 164,
                FormatCode = "$0.00"
            };
            var numberingFormats = new NumberingFormats(numberingFormat);

            var expectedValue = double.Parse(cell.CellValue.InnerText, CultureInfo.InvariantCulture).ToString(numberingFormat.FormatCode, CultureInfo.InvariantCulture);

            var result = OpenXmlHelpers.GetFormatedValue(cell, cellFormat, numberingFormats);

            result.Should().Be(expectedValue);
        }

        [Fact]
        public void Given_CellValue_And_SharedStringTable_When_SharedStringTable_Is_Null_Then_GetSharedStringItem_Returns_Null()
        {
            var cellValue = new CellValue();

            var result = OpenXmlHelpers.GetSharedStringItem(cellValue, null);

            result.Should().BeNull();
        }

        [Fact]
        public void Given_CellValue_And_SharedStringTable_When_Cell_InnerText_In_SharedStringTable_Then_GetSharedStringItem_Returns_String()
        {
            var cellValue = new CellValue(0.ToString());
            var sharedStringItem = new SharedStringItem
            {
                Text = new Text("test")
            };
            var sharedStringTable = new SharedStringTable(sharedStringItem);

            var result = OpenXmlHelpers.GetSharedStringItem(cellValue, sharedStringTable);

            result.Should().Be(sharedStringItem.InnerText);
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_When_CellFormat_Exists_Then_GetCellValueWithoutConsideringDataType_Returns_Formatted_Value()
        {
            var cell = new Cell
            {
                CellValue = new CellValue("0.01"),
                StyleIndex = 0
            };
            var cellFormat = new CellFormat
            {
                FormatId = 0,
                ApplyNumberFormat = true,
                NumberFormatId = 2
            };
            var cellFormats = new CellFormats(cellFormat);
            var stylesheet = new Stylesheet(cellFormats);

            var result = OpenXmlHelpers.GetCellValueWithoutConsideringDataType(cell, stylesheet);

            result.Should().Be("0.01");
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_When_CellFormat_Doesnt_Exist_Then_GetCellValueWithoutConsideringDataType_Returns_Value()
        {
            var cell = new Cell
            {
                CellValue = new CellValue("0.12345"),
            };
            var stylesheet = new Stylesheet();

            var result = OpenXmlHelpers.GetCellValueWithoutConsideringDataType(cell, stylesheet);

            result.Should().Be("0.12345");
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_And_SharedStringTable_When_Cell_Is_Null_Then_GetCellValue_Returns_Cell_InnerText()
        {
            var result = OpenXmlHelpers.GetCellValue(null, null, null);

            result.Should().BeNull();
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_And_SharedStringTable_When_Cell_InnerText_Is_Null_Then_GetCellValue_Returns_Cell_InnerText()
        {
            var cell = new Cell
            {
                CellValue = new CellValue(null)
            };

            var result = OpenXmlHelpers.GetCellValue(cell, null, null);

            result.Should().BeNull();
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_And_SharedStringTable_When_Cell_DataType_Is_Null_Then_GetCellValue_Returns_Cell_Without_Considering_DataType()
        {
            var cell = new Cell
            {
                CellValue = new CellValue("0.12345"),
            };
            var stylesheet = new Stylesheet();

            var result = OpenXmlHelpers.GetCellValue(cell, stylesheet, null);

            result.Should().Be("0.12345");
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_And_SharedStringTable_When_Cell_DataType_Is_String_Then_GetCellValue_Returns_Cell_CellValue_InnerText()
        {
            var cell = new Cell
            {
                CellValue = new CellValue("test"),
                DataType = CellValues.String
            };

            var result = OpenXmlHelpers.GetCellValue(cell, null, null);

            result.Should().Be("test");
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_And_SharedStringTable_When_Cell_DataType_Is_SharedString_Then_GetCellValue_Returns_SharedStringTable_SharedStringItem()
        {
            var cell = new Cell
            {
                CellValue = new CellValue(0.ToString()), // points to index 0 of shared strings table
                DataType = CellValues.SharedString
            };
            var sharedStringItem = new SharedStringItem
            {
                Text = new Text("test")
            };
            var sharedStringTable = new SharedStringTable(sharedStringItem);

            var result = OpenXmlHelpers.GetCellValue(cell, null, sharedStringTable);

            result.Should().Be("test");
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_And_SharedStringTable_When_Cell_DataType_Is_Boolean_With_Value_0_Then_GetCellValue_Returns_FALSE_String()
        {
            var cell = new Cell
            {
                CellValue = new CellValue("0"),
                DataType = CellValues.Boolean
            };

            var result = OpenXmlHelpers.GetCellValue(cell, null, null);

            result.Should().Be("FALSE");
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_And_SharedStringTable_When_Cell_DataType_Is_Boolean_With_Value_1_Then_GetCellValue_Returns_TRUE_String()
        {
            var cell = new Cell
            {
                CellValue = new CellValue("1"),
                DataType = CellValues.Boolean
            };

            var result = OpenXmlHelpers.GetCellValue(cell, null, null);

            result.Should().Be("TRUE");
        }

        [Fact]
        public void Given_Cell_And_Stylesheet_And_SharedStringTable_When_Cell_DataType_Is_Not_String_Or_Boolean_Then_GetCellValue_Returns_CellValue_InnerText()
        {
            var cell = new Cell
            {
                CellValue = new CellValue("1"),
                DataType = CellValues.Number
            };

            var result = OpenXmlHelpers.GetCellValue(cell, null, null);

            result.Should().Be("1");
        }

        [Fact]
        public void Given_Page_PageSize_XlsxSheetDimensions_OpenXmlReader_Stylesheet_And_SharedStringTable_When_Reader_Is_EOF_Then_GetRow_Returns_Null()
        {
            var page = 1;
            var pageSize = 1000;
            var dimensions = new XlsxSheetDimensions();

            var temp = Path.GetTempFileName();
            TestHelpers.Create(temp);

            using (var spreadsheetDocument = SpreadsheetDocument.Open(temp, false))
            {
                var worksheetPart = spreadsheetDocument.WorkbookPart.WorksheetParts.First();

                var reader = OpenXmlHelpers.GetOpenXmlReader(worksheetPart);

                OpenXmlHelpers.MoveReaderToFirstRow(reader); // move reader to 1st row, no rows exists so it will go to eof

                var result = OpenXmlHelpers.GetRow(page, pageSize, dimensions, reader, null, null);

                result.Should().BeNull();
            }

            File.Delete(temp);
        }

        [Fact]
        public void Given_Page_PageSize_XlsxSheetDimensions_OpenXmlReader_Stylesheet_And_SharedStringTable_When_Row_Exists_Then_GetRow_Returns_Row_As_String_List()
        {
            var page = 1;
            var pageSize = 1000;
            var dimensions = new XlsxSheetDimensions(1, 1, 1, 2);

            var temp = Path.GetTempFileName();

            var cells = new List<Cell>
            {
                new Cell { CellValue = new CellValue("1"), CellReference = "A1" },
                new Cell { CellValue = new CellValue("2"), CellReference = "B1" }
            };
            var rows = new List<Row>
            {
                new Row(cells)
            };

            TestHelpers.Create(temp, rows);

            using (var spreadsheetDocument = SpreadsheetDocument.Open(temp, false))
            {
                var worksheetPart = spreadsheetDocument.WorkbookPart.WorksheetParts.First();

                var reader = OpenXmlHelpers.GetOpenXmlReader(worksheetPart);

                OpenXmlHelpers.MoveReaderToFirstRow(reader); // move reader to 1st row

                var result = OpenXmlHelpers.GetRow(page, pageSize, dimensions, reader, null, null);

                result.Should().NotBeNull();
                result.Should().HaveCount(2);
                result[0].Should().Be("1");
                result[1].Should().Be("2");
            }

            File.Delete(temp);
        }

        [Fact]
        public void Given_Page_PageSize_XlsxSheetDimensions_OpenXmlReader_Stylesheet_And_SharedStringTable_When_Rows_Exist_Then_GetRows_Returns_Rows_As_List_Of_String_Lists()
        {
            var page = 1;
            var pageSize = 1000;
            var dimensions = new XlsxSheetDimensions(1, 2, 1, 2);

            var temp = Path.GetTempFileName();

            var rows = new List<Row>
            {
                new Row(new List<Cell>
                {
                    new Cell { CellValue = new CellValue("1"), CellReference = "A1" },
                    new Cell { CellValue = new CellValue("2"), CellReference = "B1" }
                }),
                new Row(new List<Cell>
                {
                    new Cell { CellValue = new CellValue("3"), CellReference = "A2" },
                    new Cell { CellValue = new CellValue("4"), CellReference = "B2" }
                })
            };

            TestHelpers.Create(temp, rows);

            using (var spreadsheetDocument = SpreadsheetDocument.Open(temp, false))
            {
                var worksheetPart = spreadsheetDocument.WorkbookPart.WorksheetParts.First();

                var reader = OpenXmlHelpers.GetOpenXmlReader(worksheetPart);

                OpenXmlHelpers.MoveReaderToFirstRow(reader); // move reader to 1st row

                var result = OpenXmlHelpers.GetRows(page, pageSize, dimensions, reader, null, null);

                result.Should().NotBeNull();
                result.Should().HaveCount(2);

                result[0].Should().NotBeNull();
                result[0].Should().HaveCount(2);
                result[0][0].Should().Be("1");
                result[0][1].Should().Be("2");

                result[1].Should().NotBeNull();
                result[1].Should().HaveCount(2);
                result[1][0].Should().Be("3");
                result[1][1].Should().Be("4");
            }

            File.Delete(temp);
        }

        [Fact]
        public void Given_Page_2_And_PageSize_2_XlsxSheetDimensions_With_3_Rows_When_Page_1_Is_Skipped_Then_GetRows_Returns_One_Row()
        {
            var page = 2;
            var pageSize = 2;
            var dimensions = new XlsxSheetDimensions(1, 3, 1, 2);

            var temp = Path.GetTempFileName();

            var rows = new List<Row>
            {
                new Row(new List<Cell>
                {
                    new Cell { CellValue = new CellValue("1"), CellReference = "A1" },
                    new Cell { CellValue = new CellValue("2"), CellReference = "B1" }
                }) { RowIndex = 1 },
                new Row(new List<Cell>
                {
                    new Cell { CellValue = new CellValue("3"), CellReference = "A2" },
                    new Cell { CellValue = new CellValue("4"), CellReference = "B2" }
                }) { RowIndex = 2 },
                new Row(new List<Cell>
                {
                    new Cell { CellValue = new CellValue("5"), CellReference = "A3" },
                    new Cell { CellValue = new CellValue("6"), CellReference = "B3" }
                }) { RowIndex = 3 }
            };

            TestHelpers.Create(temp, rows);

            using (var spreadsheetDocument = SpreadsheetDocument.Open(temp, false))
            {
                var worksheetPart = spreadsheetDocument.WorkbookPart.WorksheetParts.First();

                var reader = OpenXmlHelpers.GetOpenXmlReader(worksheetPart);

                OpenXmlHelpers.SkipRows(reader, (page - 1) * pageSize);

                var result = OpenXmlHelpers.GetRows(page, pageSize, dimensions, reader, null, null);

                result.Should().NotBeNull();
                result.Should().HaveCount(1);

                result[0].Should().NotBeNull();
                result[0].Should().HaveCount(2);
                result[0][0].Should().Be("5");
                result[0][1].Should().Be("6");
            }

            File.Delete(temp);
        }
    }
}
