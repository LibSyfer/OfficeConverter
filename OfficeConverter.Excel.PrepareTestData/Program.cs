using CommandLine;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Runtime.InteropServices;

internal class Program
{
    class Options
    {
        [Option('s', "source", HelpText = "Исходная директория с Excel файлами (по умолчанию - текущая)")]
        public string SourceDirectory { get; set; } = Directory.GetCurrentDirectory();

        [Option('t', "target", HelpText = "Целевая директория для XLSX файлов (по умолчанию - текущая)")]
        public string TargetDirectory { get; set; } = Directory.GetCurrentDirectory();

        [Option('f', "formats", Default = ".xlsx; .xlsm; .xlsb; .xltx; .xltm; .xlt; .xls; .ods", HelpText = "Поддерживаемые форматы (через точку с запятой)")]
        public string SupportedFormats { get; set; }

        [Option('v', "verbose", HelpText = "Подробный вывод информации")]
        public bool Verbose { get; set; } = false;

        [Option('o', "overwrite", HelpText = "Перезаписывать существующие файлы")]
        public bool Overwrite { get; set; } = false;
    }

    private static void Main(string[] args)
    {
        Parser.Default.ParseArguments<Options>(args)
                .WithParsed(RunWithOptions)
                .WithNotParsed(HandleOptionsError);
    }

    private static void HandleOptionsError(IEnumerable<CommandLine.Error> errs)
    {
        if (errs.IsVersion() || errs.IsHelp())
            return;

        Console.WriteLine("Ошибка в параметрах командной строки");
        Environment.Exit(1);
    }

    private static void RunWithOptions(Options options)
    {
        try
        {
            options.SourceDirectory = Path.GetFullPath(options.SourceDirectory);
            options.TargetDirectory = Path.GetFullPath(options.TargetDirectory);

            if (!Directory.Exists(options.SourceDirectory))
            {
                Console.WriteLine($"Исходная директория {options.SourceDirectory} не существует");
                Environment.Exit(1);
                return;
            }

            if (options.Verbose)
            {
                Console.WriteLine($"Конвертация файлов из: {options.SourceDirectory}");
                Console.WriteLine($"Сохранение результатов в: {options.TargetDirectory}");
                Console.WriteLine($"Поддерживаемые форматы: {options.SupportedFormats}");
            }

            if (!IsExcelInstalled())
            {
                Console.WriteLine("Для работы программы требуется Excel. Microsoft Excel не установлен на этом компьютере");
                Environment.Exit(1);
                return;
            }

            Application excelApp = new Application();
            excelApp.DisplayAlerts = false;
            excelApp.AskToUpdateLinks = false;
            excelApp.AlertBeforeOverwriting = false;
            #if DEBUG
            excelApp.Visible = true;
            #else
            excelApp.Visible = false;
            #endif

            var allowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlt", ".xls", ".ods"
            };

            try
            {
                ConvertAllToRandomExcelFormat(
                    targetPath: options.TargetDirectory,
                    sourcePath: options.SourceDirectory,
                    allowedExtensions: allowedExtensions,
                    excelApp: excelApp,
                    overwrite: options.Overwrite,
                    verbose: options.Verbose);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка: {ex.Message}");
            Environment.Exit(1);
        }
    }

    private static bool IsExcelInstalled()
    {
        try
        {
            using (var key = Registry.ClassesRoot.OpenSubKey("Excel.Application"))
            {
                return key != null;
            }
        }
        catch
        {
            return false;
        }
    }

    private static void ConvertAllToRandomExcelFormat(string targetPath, string sourcePath, HashSet<string> allowedExtensions, Application excelApp, bool overwrite, bool verbose)
    {
        bool hasValidFiles = false;
        var random = new Random();

        foreach (var filePath in Directory.EnumerateFiles(sourcePath, "*.*"))
        {
            if (Path.GetFileName(filePath).StartsWith("~$"))
            {
                if (verbose)
                    Console.WriteLine($"Skip temporaly file: {filePath}");
                continue;
            }
            var extension = Path.GetExtension(filePath);
            if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                var newExtension = allowedExtensions.ElementAt(random.Next(allowedExtensions.Count));
                var outputFileName = Path.GetFileNameWithoutExtension(filePath) + newExtension;
                var outputFilePath = Path.Combine(targetPath, outputFileName);
                if (File.Exists(outputFilePath))
                {
                    if (!overwrite)
                    {
                        outputFilePath = Path.Combine(targetPath, Path.GetFileNameWithoutExtension(outputFileName) + "-" + Guid.NewGuid().ToString("N") + ".xlsx");
                        if (verbose)
                            Console.WriteLine($"Перезапись выключена, будет создан новый файл: {outputFilePath}");
                    }
                    else
                    {
                        if (verbose)
                            Console.WriteLine($"Перезапись включена, будет перезаписан файл: {outputFilePath}");
                        File.Delete(outputFilePath);
                    }
                }

                if (!hasValidFiles)
                {
                    Directory.CreateDirectory(targetPath);
                    hasValidFiles = true;
                }

                ConvertToRandomExcelFormat(filePath, outputFilePath, GetFormat(newExtension), excelApp, verbose);
            }
        }

        foreach (var subdir in Directory.EnumerateDirectories(sourcePath))
        {
            var targetSubdir = Path.Combine(targetPath, Path.GetFileName(subdir));
            ConvertAllToRandomExcelFormat(targetSubdir, subdir, allowedExtensions, excelApp, overwrite, verbose);
        }
    }

    private static XlFileFormat GetFormat(string format)
    {
        return format switch
        {
            ".xlsm" => XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
            ".xlsb" => XlFileFormat.xlExcel12,
            ".xlsx" => XlFileFormat.xlOpenXMLWorkbook,
            ".xltx" => XlFileFormat.xlOpenXMLTemplate,
            ".xltm" => XlFileFormat.xlOpenXMLTemplateMacroEnabled,
            ".xlt" => XlFileFormat.xlTemplate,
            ".xls" => XlFileFormat.xlExcel8,
            ".ods" => XlFileFormat.xlOpenDocumentSpreadsheet,
            _ => throw new NotSupportedException($"Формат файла {format} не поддерживается")
        };
    }

    private static void ConvertToRandomExcelFormat(string inputFilePath, string outputPath, XlFileFormat format, Application excelApp, bool verbose)
    {
        Workbook? workbook = null;
        try
        {
            workbook = excelApp.Workbooks.Open(inputFilePath);
            workbook.SaveAs(
                Filename: outputPath,
                FileFormat: format,
                ConflictResolution: XlSaveConflictResolution.xlUserResolution);
            if (verbose)
                Console.WriteLine($"Конвертирован: {inputFilePath} -> {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при конвертации {inputFilePath}: {ex.Message}");
        }
        finally
        {
            if (workbook != null)
            {
                workbook.Close();
                Marshal.FinalReleaseComObject(workbook);
                workbook = null;
            }
        }
    }
}