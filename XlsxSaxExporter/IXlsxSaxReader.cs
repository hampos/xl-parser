using System;
using System.Collections.Generic;

namespace XlsxSaxExporter
{
    public interface IXlsxSaxReader : IDisposable
    {
        XlsxSheetDimensions Dimensions { get; }
        List<List<string>> Read(int page);
    }
}