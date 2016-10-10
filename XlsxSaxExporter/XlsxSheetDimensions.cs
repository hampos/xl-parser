namespace XlsxSaxExporter
{
    public class XlsxSheetDimensions
    {
        public XlsxSheetDimensions()
        { }

        public XlsxSheetDimensions(int minRowNum, int maxRowNum, int minColNum, int maxColNum)
        {
            MinRowNum = minRowNum;
            MaxRowNum = maxRowNum;
            MinColNum = minColNum;
            MaxColNum = maxColNum;
        }

        public int MinRowNum { get; private set; }
        public int MaxRowNum { get; private set; }
        public int MinColNum { get; private set; }
        public int MaxColNum { get; private set; }
    }
}
