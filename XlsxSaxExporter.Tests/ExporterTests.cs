using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace XlsxSaxExporter.Tests
{
    public class ExporterTests
    {
        [Fact]
        public void Given_Xlsx_Path_I_Can_Export_All_Rows()
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
            TestHelpers.Create(temp, rows, "A1:C2");

            var result = Exporter.Export(temp);

            result.Should().HaveCount(3);

            File.Delete(temp);
        }

        [Fact]
        public void Given_Xlsx_Path_And_Internal_PageSize_I_Can_Export_All_Rows()
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

            var result = Exporter.Export(temp, 1);

            result.Should().HaveCount(3);

            File.Delete(temp);
        }
    }
}
