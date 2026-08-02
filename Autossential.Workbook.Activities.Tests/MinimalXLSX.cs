using System.IO.Compression;
using System.Text;

namespace Autossential.Workbook.Activities.Tests
{
    public static class MinimalXLSX
    {
        public static string Create(int cols, int rows)
        {
            var headers = Enumerable.Range(0, cols).Select(i => $"Col{i + 1}").ToArray();
            var data = new List<string[]>();
            for (int i = 0; i < rows; i++)
            {
                data.Add(Enumerable.Range(0, cols).Select(j => $"C{j + 1}R{i + 1}").ToArray());
            }
            var path = Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");
            //path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", Path.GetFileName(path));
            CreateMinimalXlsx(path, headers, data.ToArray());
            return path;
        }

        static void CreateMinimalXlsx(string path, string[] headers, string[][] rows)
        {
            File.Delete(path);
            using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

            WriteEntry(archive, "[Content_Types].xml", ContentTypesXml);
            WriteEntry(archive, "_rels/.rels", RelsXml);
            WriteEntry(archive, "xl/workbook.xml", WorkbookXml);
            WriteEntry(archive, "xl/_rels/workbook.xml.rels", WorkbookRelsXml);
            WriteEntry(archive, "xl/styles.xml", StylesXml);
            WriteEntry(archive, "xl/worksheets/sheet1.xml", BuildSheetXml(headers, rows));
        }

        static string BuildSheetXml(string[] headers, string[][] rows)
        {
            var sb = new StringBuilder();
            sb.Append("""<?xml version="1.0" encoding="utf-8"?>""");
            sb.Append("""<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""");
            sb.Append("<x:sheetData>");

            // Header row - sem r em <row> nem em <c>
            sb.Append("<x:row>");
            foreach (var h in headers)
                sb.Append($"""<x:c t="inlineStr"><x:is><x:t>{System.Security.SecurityElement.Escape(h)}</x:t></x:is></x:c>""");
            sb.Append("</x:row>");

            // Data rows
            foreach (var row in rows)
            {
                sb.Append("<x:row>");
                foreach (var val in row)
                    sb.Append($"""<x:c t="inlineStr"><x:is><x:t>{System.Security.SecurityElement.Escape(val)}</x:t></x:is></x:c>""");
                sb.Append("</x:row>");
            }

            sb.Append("</x:sheetData></x:worksheet>");
            return sb.ToString();
        }

        static void WriteEntry(ZipArchive archive, string name, string content)
        {
            var entry = archive.CreateEntry(name);
            using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
            writer.Write(content);
        }

        const string ContentTypesXml = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
""";

        const string RelsXml = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
""";

        const string WorkbookXml = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>
""";

        const string WorkbookRelsXml = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
""";

        const string StylesXml = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="1"><fill/></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf/></cellXfs>
</styleSheet>
""";
    }
}
