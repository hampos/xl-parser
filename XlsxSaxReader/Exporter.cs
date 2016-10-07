using System.Collections.Generic;

namespace XlsxSaxReader
{
    public class Exporter
    {
        public static List<List<string>> Export(string path, int pageSize)
        {
            IXlsxSaxReader xlsxSaxReader = new XlsxSaxReader(path, pageSize);

            int page = 1;
            var rows = new List<List<string>>(xlsxSaxReader.Dimensions.MaxRowNum);

            do
            {
                var result = xlsxSaxReader.Read(page++);
                if (result.Count == 0)
                    break;

                rows.AddRange(result);
            } while (true);

            return rows;
        }
    }
}
