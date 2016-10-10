using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSaxExporter
{
    public class XlsxSaxReader : IXlsxSaxReader
    {
        private readonly string _path;
        private readonly int _pageSize;

        private SpreadsheetDocument _spreadsheetDoc;
        private WorksheetPart _worksheetPart;

        private OpenXmlReader _openXmlReader;
        private int _nextPageNum = 1;

        public XlsxSaxReader(string path, int pageSize = 1000)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentNullException("path");
            if (pageSize < 1) throw new ArgumentException("PageSize must be a positive value");

            _path = path;
            _pageSize = pageSize;

            Setup(path);
        }

        public XlsxSheetDimensions Dimensions { get; private set; }

        public IList<IList<string>> Read(int page, int pageSize = 0)
        {
            Setup(_path, page);

            var rows = OpenXmlHelpers.GetRows(
                page,
                pageSize == 0 ? _pageSize : pageSize,
                Dimensions,
                _openXmlReader,
                _spreadsheetDoc.WorkbookPart.WorkbookStylesPart == null ? null : _spreadsheetDoc.WorkbookPart.WorkbookStylesPart.Stylesheet,
                _spreadsheetDoc.WorkbookPart.SharedStringTablePart == null ? null : _spreadsheetDoc.WorkbookPart.SharedStringTablePart.SharedStringTable);

            _nextPageNum = page + 1;

            return rows;
        }

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

        private void Setup(string path, int page = 0)
        {
            if (page == 0)
            {
                Dimensions = OpenXmlHelpers.GetDimensions(path);
                return;
            }

            if (page > 1 && page == _nextPageNum)
            {
                return;
            }

            DisposeReader();
            DisposeDocument();

            _spreadsheetDoc = SpreadsheetDocument.Open(path, false);
            var sheet = _spreadsheetDoc.WorkbookPart.Workbook.Descendants<Sheet>().First();
            _worksheetPart = (WorksheetPart)_spreadsheetDoc.WorkbookPart.GetPartById(sheet.Id);

            _openXmlReader = OpenXmlReader.Create(_worksheetPart);
            OpenXmlHelpers.SkipRows(_openXmlReader, (page - 1) * _pageSize);
        }

        private void DisposeReader()
        {
            if (_openXmlReader != null)
            {
                _openXmlReader.Close();
                _openXmlReader.Dispose();
            }
        }

        private void DisposeDocument()
        {
            if (_spreadsheetDoc != null)
            {
                _spreadsheetDoc.Close();
                _spreadsheetDoc.Dispose();
            }
        }
    }
}
