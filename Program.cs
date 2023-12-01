namespace excelToJson
{
    using ClosedXML.Excel;
    using DocumentFormat.OpenXml.Packaging;
    using Newtonsoft.Json;
    using System.Collections.Generic;
    using System.Xml;

    class Program
    {
        static void Main(string[] args)
        {
            var workbookPath = "/Users/bulent.erdem/Desktop/TezcanİşKazaAnalizi.xlsx"; // Excel dosyasının yolu
            var jsonString = ConvertExcelToJson(workbookPath);
            System.Console.WriteLine(jsonString);
        }

        static string ConvertExcelToJson(string filePath)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1); // Birinci sayfa
            var rows = worksheet.RangeUsed().RowsUsed();

            // Anahtarları içeren ilk satırı alın
            var headers = rows.First().Cells().Select(c => c.GetValue<string>().Trim());

            var data = new List<Dictionary<string, object>>();

            // İkinci satırdan itibaren verileri işleyin
            foreach (var row in rows.Skip(1))
            {
                var rowData = new Dictionary<string, object>();
                foreach (var cell in row.Cells())
                {
                    var key = headers.ElementAt(cell.Address.ColumnNumber - 1);
                    var value = cell.GetValue<string>().Trim();
                    rowData[key] = value;
                }
                data.Add(rowData);
            }

            var jsonOutputPath = "/Users/bulent.erdem/Desktop/analiziJson.json"; // JSON dosyasının kaydedileceği yol
            // JSON string'ini dosyaya yaz
            File.WriteAllText(jsonOutputPath, JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented));

            return JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
        }
    }

}