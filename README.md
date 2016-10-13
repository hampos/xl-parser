# xl-parser
A blazing fast SAX reader for xlsx files with paging and low memory consumption

The exporter uses `OpenXml SDK 2.5` to export xlsx data with paging support.
Page size is required so the reader can work on a fixed page size while exporting.
The result is a `List<List<string>>` to offer an easy way to enumerate the data.

The exporter was built to be used on a server environment for simple exporting purposes, 
where it wasn't possible to use a driver and to overcome out-of-memory situations in case of large spreadsheets.

Current version supports only exporting the first sheet by default. 

Example usage (see `ConsoleApp` as well)

```
using XlsxSaxExporter;

var path = "c:\\test.xlsx";

var rows = Exporter.Export(path);

return rows;
```

If the need exists to manipulate page data as they come, then use `XlsxSaxReader` and process page data accordingly.
The example shown below is the same as the `Exporter.Export` method:

```
using XlsxSaxExporter;

var path = "c:\\test.xlsx";
var pageSize = 1000;

IXlsxSaxReader xlsxSaxReader = new XlsxSaxReader(path, pageSize);

int page = 1;
var rows = new List<IList<string>>(xlsxSaxReader.Dimensions.MaxRowNum);

do
{
    var result = xlsxSaxReader.Read(page++);
    if (result.Count == 0)
        break;
        
    // Process row data here

    rows.AddRange(result);
} while (true);

return rows;
```

Make sure to add a reference to `WindowsBase` since it is required by `OpenXml SDK`, the code doesn't compile without it.

Feel free to send your feedback or fork the project.

**Important notice**

Since `OpenXmlReader` is thread safe ([MSDN](https://msdn.microsoft.com/en-us/library/documentformat.openxml.openxmlreader(v=office.15).aspx)) only when declared as a `public static` property. As a result, `XlsxSaxReader` is not guarranteed to be thread safe.

Moreover, at the time of this writing, `Open XML SDK` is already at version `2.6`, as found [here](https://github.com/OfficeDev/Open-XML-SDK). This version solves known issues with `System.IO.Packaging` from `WindowsBase`. Unfortunately, no nuget package exists yet, so it is recommended to build those sources and then add references to `DocumentFormat.OpenXml.dll` and `System.IO.Packaging.dll`. If you follow this road, make sure to remove the sdk from nuget packages as well as the referenced `WindowsBase` library.
