﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSaxReader
{
    public class XlsxSaxReader : IXlsxSaxReader
    {
        private readonly string _path;
        private readonly int _pageSize;

        private SpreadsheetDocument _spreadsheetDoc;
        private WorksheetPart _worksheetPart;

        private OpenXmlReader _openXmlReader;
        private int _nextPageNum = 1;

        public XlsxSaxReader(string path, int pageSize)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentNullException("path");
            if (pageSize < 1) throw new ArgumentException("PageSize must be a positive value");

            _path = path;
            _pageSize = pageSize;

            Setup(path);
        }

        public XlsxSheetDimensions Dimensions { get; private set; }

        public List<List<string>> Read(int page)
        {
            if (page * _pageSize > Dimensions.MaxRowNum + _pageSize)
                return new List<List<string>>();

            Setup(_path, page);

            var rows = OpenXmlHelpers.GetRows(
                page,
                _pageSize,
                Dimensions,
                _openXmlReader,
                _spreadsheetDoc.WorkbookPart.WorkbookStylesPart.Stylesheet,
                _spreadsheetDoc.WorkbookPart.SharedStringTablePart.SharedStringTable);

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
            OpenXmlHelpers.SkipRows(_openXmlReader, page, _pageSize);
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
