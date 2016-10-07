using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using System.Collections.Generic;
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
            Create(temp);

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

            var cell = new Cell()
            {
                CellValue = new CellValue("1")
            };
            var sheetData = new SheetData(new Row(cell));
            
            Create(temp, sheetDimensionRef: "A1:B2");

            var dimensions = OpenXmlHelpers.GetDimensions(temp);

            dimensions.Should().NotBeNull();
            dimensions.MinRowNum.Should().Be(1);
            dimensions.MaxRowNum.Should().Be(2);
            dimensions.MinColNum.Should().Be(1);
            dimensions.MaxColNum.Should().Be(2);

            File.Delete(temp);
        }

        private void Create(string filepath, IEnumerable<Row> rows = null, string sheetDimensionRef = null)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();

            // Add SheedData
            SheetData sheetData = new SheetData();
            if (rows != null)
                sheetData.Append(rows);

            worksheetPart.Worksheet = new Worksheet(sheetData);

            // Add a SheetDimension
            SheetDimension sheetDimension = new SheetDimension() { Reference = sheetDimensionRef };
            worksheetPart.Worksheet.Append(sheetDimension);

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
        }
    }
}
