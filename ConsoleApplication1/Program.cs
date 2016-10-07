using System;
using System.Collections.Generic;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = null;
            if (args.Length > 0)
            {
                path = args[0];
            }

            while(string.IsNullOrWhiteSpace(path))
            {
                Console.WriteLine("Enter the xlsx file path: ");
                path = Console.ReadLine();
            }

            var xlsxSaxReader = new XlsxSaxReader.XlsxSaxReader(path, 1000);

            Console.WriteLine();
            Console.WriteLine("Opened XLSX found at: " + path);
            Console.WriteLine("First row: " + xlsxSaxReader.Dimensions.MinRowNum);
            Console.WriteLine("Last row: " + xlsxSaxReader.Dimensions.MaxRowNum);
            Console.WriteLine("First column: " + xlsxSaxReader.Dimensions.MinColNum);
            Console.WriteLine("Last column: " + xlsxSaxReader.Dimensions.MaxColNum);
            Console.WriteLine("Reading...");
            Console.WriteLine();

            int page = 1;
            var rows = new List<List<string>>();

            try
            {
                do
                {
                    var result = xlsxSaxReader.Read(page++);
                    if (result.Count == 0)
                        break;

                    rows.AddRange(result);
                } while (true);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                Console.WriteLine(page);
            }

            Console.WriteLine("Read " + rows.Count + " rows in total.");
            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadLine();
        }
    }
}
