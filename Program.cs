namespace excelToJson
{
    using ClosedXML.Excel;
    using DocumentFormat.OpenXml.Packaging;
    using Newtonsoft.Json;
    using System.Collections.Generic;
    using System.Xml;
    //C:\Users\bulent.erdem\Desktop\excelToJson\tezcanlarİş KazasıAnalizi\TezcanİşKazaAnalizi.xlsx
    class Program
    {
        static void Main(string[] args)
        {
            var workbookPath = "C:\\Users\\bulent.erdem\\Desktop\\excelToJson\\tezcanlarİş KazasıAnalizi\\TezcanİşKazaAnalizi.xlsx"; // Excel dosyasının yolu
            var jsonString = ConvertExcelToJson(workbookPath);
            var jsonOutputPath = "C:\\Users\\bulent.erdem\\Desktop\\excelToJson\\tezcanlarİş KazasıAnalizi\\analiziJson.json"; // JSON dosyasının kaydedileceği yol
            File.WriteAllText(jsonOutputPath, jsonString);
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

            return JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
        }


        static string ConvertExcelToJson2(string filePath)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed();

            // Sadece belirli sütun başlıklarını almak için burada bir liste tanımlayın
            var desiredHeaders = new List<string> { "İstenenSütun1", "İstenenSütun2", "İstenenSütun3" }; // Örnek sütun adları

            // Anahtarları içeren ilk satırı alın ve yalnızca istenen sütunları filtreleyin
            var headers = rows.First().Cells()
                .Where(c => desiredHeaders.Contains(c.GetValue<string>().Trim()))
                .Select(c => c.GetValue<string>().Trim());

            var data = new List<Dictionary<string, object>>();

            // İkinci satırdan itibaren verileri işleyin
            foreach (var row in rows.Skip(1))
            {
                var rowData = new Dictionary<string, object>();
                foreach (var cell in row.CellsUsed())
                {
                    var header = cell.WorksheetColumn().ColumnLetter();
                    var key = worksheet.Cell(1, header).GetValue<string>().Trim();
                    if (desiredHeaders.Contains(key))
                    {
                        var value = cell.GetValue<string>().Trim();
                        rowData[key] = value;
                    }
                }
                if (rowData.Count > 0) // Yalnızca istenen sütunlardan veri içeriyorsa ekle
                {
                    data.Add(rowData);
                }
            }

            return JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
        }


    }

}