using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

namespace XlsxSaxReader
{
    public class XlsxSaxReader
    {
        private readonly string _path;
        private readonly SharedStringTable _sharedStringTable;
        private readonly SpreadsheetDocument _spreadsheetDoc;
        private readonly Worksheet _worksheet;
        private readonly OpenXmlReader _openXmlReader;
        private readonly Dictionary<uint, string> _cellFormats;
        private int lastRequestPage = 0;

        public XlsxSaxReader(string path)
        {
            _path = path;

            _spreadsheetDoc = SpreadsheetDocument.Open(path, false);

            var sheet = _spreadsheetDoc.WorkbookPart.Workbook.Descendants<Sheet>().First();
            var worksheetPart = (WorksheetPart)_spreadsheetDoc.WorkbookPart.GetPartById(sheet.Id);
            _openXmlReader = OpenXmlHelpers.GetOpenXmlReader(worksheetPart);
            _cellFormats = OpenXmlHelpers.GetCellFormats(_spreadsheetDoc);
            Dimensions = OpenXmlHelpers.GetDimensions(_openXmlReader);
        }

        public XlsxSheetDimensions Dimensions { get; private set; }

        public void Dispose()
        {
            if (_openXmlReader != null)
            {
                _openXmlReader.Close();
                _openXmlReader.Dispose();
            }

            if (_spreadsheetDoc != null)
            {
                _spreadsheetDoc.Close();
                _spreadsheetDoc.Dispose();
            }
        }
    }
}
