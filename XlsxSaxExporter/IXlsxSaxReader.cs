using System.Collections.Generic;

namespace XlsxSaxExporter
{
    public interface IXlsxSaxReader
    {
        XlsxSheetDimensions Dimensions { get; }

        void Dispose();
        List<List<string>> Read(int page);
    }
}