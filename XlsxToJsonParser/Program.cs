using Newtonsoft.Json;
using OfficeOpenXml;

namespace XlsxToJsonParser
{
    internal class Program
    {
        public class HeadCount
        {
            [JsonProperty("date")]
            public required string Date { get; set; }
            [JsonProperty("number")]
            public required int Number { get; set; }
        }

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage("Тестовое задание.xlsx");
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Состав на последний день месяца"];
            string address = worksheet.DimensionByValue.Address;
            ExcelRange cells = worksheet.Cells[address];
            if (cells.Rows == 0)
            {
                Console.WriteLine("Sheet is empty");
                return;
            }

            if (cells.Rows == 1)
            {
                Console.WriteLine("No data provided");
                return;
            }

            string[] titles = [cells.TakeSingleCell(0, 0).Text, cells.TakeSingleCell(0, 1).Text];
            ExcelRangeBase dataCells = cells.SkipRows(1);
            List<HeadCount> data = [];
            for (int i = 0; i < dataCells.Rows; i++)
            {
                ExcelRangeBase row = dataCells.TakeSingleRow(i);
                ExcelRangeBase cell = row.TakeSingleColumn(0);

                bool parseResult = DateOnly.TryParse(cell.Text, out DateOnly resultDate);
                if (parseResult == false)
                {
                    Console.WriteLine($"Date parse error in {cell.Address}");
                    continue;
                }

                string date = resultDate.ToString("o");

                cell = row.TakeSingleColumn(1);
                parseResult = int.TryParse(cell.Text, out int resultNumber);
                if (parseResult == false)
                {
                    Console.WriteLine($"Number parse error in {cell.Address}");
                    continue;
                }

                data.Add(new HeadCount { Date = date, Number = resultNumber });
            }

            List<object> objects = data.ToList<object>().Prepend(titles).ToList();
            string serializedObject = JsonConvert.SerializeObject(objects, Formatting.Indented);
            string newFilePath = Path.ChangeExtension(package.File.FullName, "json");
            File.WriteAllText(newFilePath, serializedObject);
        }
    }
}
