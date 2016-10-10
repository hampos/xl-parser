using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace XlsxSaxExporter.Tests
{
    public class XlsxSaxReaderTests
    {
        [Fact]
        public void Given_Xlsx_File_Path_And_PageSize_I_Can_Create_XlsxSaxReader_With_XlsxSheetDimensions()
        {
            var temp = Path.GetTempFileName();
            TestHelpers.Create(temp, sheetDimensionRef: "A1:B2");
            var pageSize = 1000;

            using (var reader = new XlsxSaxReader(temp, pageSize))
            {
                reader.Should().NotBeNull();
                reader.Dimensions.Should().NotBeNull();
            };
            
            File.Delete(temp);
        }

        [Fact]
        public void Given_Null_File_Path_When_Creating_Then_It_Throws()
        {
            var exc = Assert.Throws<ArgumentNullException>(() => new XlsxSaxReader(null));

            exc.ParamName.Should().Be("path");
        }

        [Fact]
        public void Given_Less_Than_One_PageSize_When_Creating_Then_It_Throws()
        {
            Assert.Throws<ArgumentException>(() => new XlsxSaxReader("path", 0));
        }

        [Fact]
        public void Given_Page_Then_Read_Returns_Rows()
        {
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
            TestHelpers.Create(temp, rows, "A1:B3");
            var pageSize = 1000;

            using (var reader = new XlsxSaxReader(temp, pageSize))
            {
                var result = reader.Read(1);

                result.Should().HaveCount(3);
            };

            File.Delete(temp);
        }

        [Fact]
        public void Given_PageSize_Smaller_Than_Rows_Dimension_Then_Read_Returns_Rows_Of_Page()
        {
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
            TestHelpers.Create(temp, rows, "A1:B3");
            var pageSize = 2;

            using (var reader = new XlsxSaxReader(temp, pageSize))
            {
                var result = reader.Read(1);

                result.Should().HaveCount(2);

                result = reader.Read(2);

                result.Should().HaveCount(1);
            };

            File.Delete(temp);
        }

        [Fact]
        public void Given_Next_Page_Requested_Not_Same_As_Expected_Page_Then_Read_Returns_Requested_Page()
        {
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
            TestHelpers.Create(temp, rows, "A1:B3");
            var pageSize = 2;

            using (var reader = new XlsxSaxReader(temp, pageSize))
            {
                var result = reader.Read(2);

                result.Should().HaveCount(1);
            };

            File.Delete(temp);
        }

        [Fact]
        public void Given_Next_Page_Requested_Larger_Than_Total_Pages_Then_Read_Returns_Empty_Result()
        {
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
            TestHelpers.Create(temp, rows, "A1:B3");
            var pageSize = 2;

            using (var reader = new XlsxSaxReader(temp, pageSize))
            {
                var result = reader.Read(3);

                result.Should().BeEmpty();
            };

            File.Delete(temp);
        }
    }
}
