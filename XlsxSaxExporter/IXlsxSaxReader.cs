using System;
using System.Collections.Generic;

namespace XlsxSaxExporter
{
    public interface IXlsxSaxReader : IDisposable
    {
        XlsxSheetDimensions Dimensions { get; }
        IList<IList<string>> Read(int page, int pageSize = 0);
    }
}