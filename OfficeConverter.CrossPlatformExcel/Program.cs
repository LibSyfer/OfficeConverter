using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeConverter.CrossPlatformExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            const string INPUT_DATA = "D:\\Data";
            const string OUTPUT_FOLDER = "D:\\Output";

            if (!Directory.Exists(OUTPUT_FOLDER))
                Directory.CreateDirectory(OUTPUT_FOLDER);

            if (!Directory.Exists(INPUT_DATA))
                Directory.CreateDirectory(INPUT_DATA);

            foreach (var file in Directory.GetFiles(INPUT_DATA, "*.*"))
            {
                try
                {
                    ConvertToXlsx(file, OUTPUT_FOLDER);
                }
                catch (NotSupportedException ex)
                {
                    Console.WriteLine($"Формат не поддерживается: {ex}");
                }
                //catch (Exception ex)
                //{
                //    Console.WriteLine($"Ошибка выполнения: {ex}");
                //}
            }
        }

        private static void ConvertToXlsx(string inputFilePath, string outputFolder)
        {
            var extension = Path.GetExtension(inputFilePath).ToLower();
            var outputFileName = Path.GetFileNameWithoutExtension(inputFilePath) + ".xlsx";
            var outputFilePath = Path.Combine(outputFolder, outputFileName);

            if (new[] { ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm" }.Any(pattern => pattern == extension))
            {
                ConvertWithOpenXmlSdk(inputFilePath, outputFilePath);
            }
            else
            {
                throw new NotSupportedException($"Format {extension} not supported");
            }
        }

        private static void ConvertWithOpenXmlSdk(string inputFilePath, string outputFilePath)
        {
            using (var newWorkbook = SpreadsheetDocument.Create(outputFilePath, SpreadsheetDocumentType.Workbook))
            {
                using (var workbook = SpreadsheetDocument.Open(inputFilePath, false))
                {
                    var workbookPart = workbook.WorkbookPart;
                    if (workbookPart != null)
                    {
                        var newWorkbookPart = newWorkbook.AddWorkbookPart();
                        newWorkbookPart.Workbook = (Workbook)workbookPart.Workbook.CloneNode(true);

                        var workbookStylesPart = workbookPart.WorkbookStylesPart;
                        if (workbookStylesPart != null)
                        {
                            var newStylesPart = newWorkbookPart.AddNewPart<WorkbookStylesPart>();
                            newStylesPart.Stylesheet = (Stylesheet)workbookStylesPart.Stylesheet.CloneNode(true);
                        }
                        else
                        {
                            Console.WriteLine("Null workbook styles part");
                        }

                        foreach (var sheet in workbookPart.WorksheetParts)
                        {
                            newWorkbookPart.AddPart(sheet);
                        }

                        if (workbookPart.VbaProjectPart != null)
                        {
                            newWorkbookPart.AddPart(workbookPart.VbaProjectPart);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Null workbook part");
                    }
                }
            }
        }
    }
}
