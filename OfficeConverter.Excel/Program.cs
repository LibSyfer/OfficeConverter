using CommandLine;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Runtime.InteropServices;

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

internal class Program
{
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
                ConvertAllToXlsx(
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

    private static void ConvertAllToXlsx(string targetPath, string sourcePath, HashSet<string> allowedExtensions, Application excelApp, bool overwrite, bool verbose)
    {
        bool hasValidFiles = false;

        foreach (var filePath in Directory.EnumerateFiles(sourcePath, "*.*"))
        {
            if (Path.GetFileName(filePath).StartsWith("~$"))
            {
                if (verbose)
                    Console.WriteLine($"Skip temporaly file: {filePath}");
                continue;
            }
            var extension = Path.GetExtension(filePath);
            if (allowedExtensions.Contains(extension))
            {
                Thread.Sleep(500);

                var outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".xlsx";
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

                ConvertToXlsx(filePath, outputFilePath, excelApp, verbose);
            }
        }

        foreach (var subdir in Directory.EnumerateDirectories(sourcePath))
        {
            var targetSubdir = Path.Combine(targetPath, Path.GetFileName(subdir));
            ConvertAllToXlsx(targetSubdir, subdir, allowedExtensions, excelApp, overwrite, verbose);
        }
    }

    private static void ConvertToXlsx(string inputFilePath, string outputPath, Application excelApp, bool verbose)
    {
        Workbook? workbook = null;
        try
        {
            workbook = excelApp.Workbooks.Open(inputFilePath);
            workbook.SaveAs(outputPath, XlFileFormat.xlOpenXMLWorkbook);
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