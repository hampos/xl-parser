using System.Collections.Generic;

namespace XlsxSaxReader
{
    internal class OpenXmlConstants
    {
        static OpenXmlConstants()
        {
            DefaultNumberingFormats = new Dictionary<uint, string>
            {
                { 0, "General" },
                { 1, "0" },
                { 2, "0.00" },
                { 3, "#,##0" },
                { 4, "#,##0.00" },
                { 9, "0%" },
                { 10, "0.00%" },
                { 11, "0.00E+00" },
                { 12, "# ?/?" },
                { 13, "# ??/??" },
                { 14, "d/M/yyyy" },
                { 15, "d-MMM-yy" },
                { 16, "d-MMM" },
                { 17, "MMM-yy" },
                { 18, "h:mm tt" },
                { 19, "h:mm:ss tt" },
                { 20, "H:mm" },
                { 21, "H:mm:ss" },
                { 22, "M/d/yyyy H:mm" },
                { 37, "#,##0 ;(#,##0)" },
                { 38, "#,##0 ;[Red](#,##0)" },
                { 39, "#,##0.00;(#,##0.00)" },
                { 40, "#,##0.00;[Red](#,##0.00)" },
                { 45, "mm:ss" },
                { 46, "[h]:mm:ss" },
                { 47, "mmss.0" },
                { 48, "##0.0E+0" },
                { 49, "@" }
            };

            DateTimeNumberingFormats = new List<uint>
            {
                14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
            };
        }

        internal static Dictionary<uint, string> DefaultNumberingFormats;
        internal static List<uint> DateTimeNumberingFormats;
    }
}
