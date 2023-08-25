using Aspose.Cells;
using Aspose.Cells.Utility;

namespace JsonToXlsx
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            string workingDirectory = Environment.CurrentDirectory;
            //Path
            string filePath = Directory.GetParent(workingDirectory).Parent.Parent.FullName;

            filePath = Path.Combine(filePath, "./Info/Data.json");

            // Read JSON file
            string jsonInput = File.ReadAllText(filePath);

            // Set json layout options
            JsonLayoutOptions options = new JsonLayoutOptions();
            options.ArrayAsTable = true;

            JsonUtility.ImportData(jsonInput, sheet.Cells, 0, 0, options);

            workbook.Save("importedData.xlsx");

            //Console.WriteLine("Hello, World!");
        }
    }
}