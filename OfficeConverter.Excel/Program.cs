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

    [Option('l', "log", HelpText = "Логгирование в файл \"log.txt\"")]
    public bool LogInFile { get; set; } = false;
}

internal class Program
{
    private static string LogFilePath = Path.Combine(Directory.GetCurrentDirectory(), "log.txt");
    private static string ErrorLogFilePath = Path.Combine(Directory.GetCurrentDirectory(), "errorLog.txt");

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
                if (options.LogInFile)
                    File.AppendAllText(LogFilePath, $"Исходная директория {options.SourceDirectory} не существует");
                Environment.Exit(1);
                return;
            }

            if (options.Verbose)
            {
                Console.WriteLine($"Конвертация файлов из: {options.SourceDirectory}");
                Console.WriteLine($"Сохранение результатов в: {options.TargetDirectory}");
                Console.WriteLine($"Поддерживаемые форматы: {options.SupportedFormats}");
            }
            if (options.LogInFile)
            {
                File.AppendAllText(LogFilePath, $"Конвертация файлов из: {options.SourceDirectory}\n");
                File.AppendAllText(LogFilePath, $"Сохранение результатов в: {options.TargetDirectory}\n");
                File.AppendAllText(LogFilePath, $"Поддерживаемые форматы: {options.SupportedFormats}\n");
            }

            if (!IsExcelInstalled())
            {
                Console.WriteLine("Для работы программы требуется Excel. Microsoft Excel не установлен на этом компьютере");
                if (options.LogInFile)
                    File.AppendAllText(LogFilePath, "Для работы программы требуется Excel. Microsoft Excel не установлен на этом компьютере\n");
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
                if (!CanWriteToFolder(options.TargetDirectory))
                {
                    Console.WriteLine($"Недостаточно прав для создания файлов в директории {options.TargetDirectory}");
                    if (options.LogInFile)
                        File.AppendAllText(LogFilePath, $"Недостаточно прав для создания файлов в директории {options.TargetDirectory}\n");
                    Environment.Exit(1);
                    return;
                }

                ConvertAllToXlsx(
                    targetPath: options.TargetDirectory,
                    sourcePath: options.SourceDirectory,
                    allowedExtensions: allowedExtensions,
                    excelApp: excelApp,
                    overwrite: options.Overwrite,
                    verbose: options.Verbose,
                    logInFile: options.LogInFile);
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
            if (options.LogInFile)
                File.AppendAllText(LogFilePath, $"Ошибка: {ex.Message}\n");
            File.AppendAllText(ErrorLogFilePath, $"[{DateTime.UtcNow}] Глобальная ошибка:\n{ex}\n");
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

    private static bool CanWriteToFolder(string folderPath)
    {
        try
        {
            string tempFile = Path.Combine(folderPath, Guid.NewGuid().ToString() + ".tmp");
            File.WriteAllText(tempFile, "test");
            File.Delete(tempFile);
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    private static void ConvertAllToXlsx(string targetPath, string sourcePath, HashSet<string> allowedExtensions, Application excelApp, bool overwrite, bool verbose, bool logInFile)
    {
        bool hasValidFiles = false;

        foreach (var filePath in Directory.EnumerateFiles(sourcePath, "*.*"))
        {
            if (Path.GetFileName(filePath).StartsWith("~$"))
            {
                if (verbose)
                    Console.WriteLine($"Пропуск временного файла excel: {filePath}");
                if (logInFile)
                    File.AppendAllText(LogFilePath, $"Пропуск временного файла excel: {filePath}\n");
                continue;
            }
            var extension = Path.GetExtension(filePath);
            if (allowedExtensions.Contains(extension))
            {
                var outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".xlsx";
                var outputFilePath = Path.Combine(targetPath, outputFileName);
                if (File.Exists(outputFilePath))
                {
                    if (!overwrite)
                    {
                        outputFilePath = Path.Combine(targetPath, Path.GetFileNameWithoutExtension(outputFileName) + "-" + Guid.NewGuid().ToString("N") + ".xlsx");
                        if (verbose)
                            Console.WriteLine($"Перезапись выключена, будет создан новый файл: {outputFilePath}");
                        if (logInFile)
                            File.AppendAllText(LogFilePath, $"Перезапись выключена, будет создан новый файл: {outputFilePath}\n");
                    }
                    }

                if (!hasValidFiles)
                {
                    Directory.CreateDirectory(targetPath);
                    hasValidFiles = true;
                }

                var result = ConvertToXlsx(filePath, outputFilePath, excelApp, verbose, logInFile);
                if (result && overwrite)
                {
                    File.Delete(filePath);
                    if (verbose)
                        Console.WriteLine($"Перезапись включена, файл {filePath} удален");
                    if (logInFile)
                        File.AppendAllText(LogFilePath, $"Перезапись включена, файл {filePath} удален\n");
            }
        }
        }

        foreach (var subdir in Directory.EnumerateDirectories(sourcePath))
        {
            var targetSubdir = Path.Combine(targetPath, Path.GetFileName(subdir));
            ConvertAllToXlsx(targetSubdir, subdir, allowedExtensions, excelApp, overwrite, verbose, logInFile);
        }
    }

    private static bool ConvertToXlsx(string inputFilePath, string outputPath, Application excelApp, bool verbose, bool logInFile)
    {
        Workbook? workbook = null;

        try
        {
            if (!File.Exists(inputFilePath))
            {
                throw new FileNotFoundException($"Исходный файл не найден: {inputFilePath}");
            }
            if (outputPath.Length > 260)
            {
                throw new PathTooLongException($"Слишком длинный путь к файлу: {outputPath}");
            }

            workbook = TryOpenWorkbook(inputFilePath, excelApp, 6, verbose, logInFile);
            if (workbook is not null)
            {
            workbook.SaveAs(
                Filename: outputPath,
                FileFormat: XlFileFormat.xlOpenXMLWorkbook,
                ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges,
                Local: true,
                    AddToMru: false
                    );
            if (verbose)
                Console.WriteLine($"Конвертирован: {inputFilePath} -> {outputPath}");
            if (logInFile)
                File.AppendAllText(LogFilePath, $"Конвертирован: {inputFilePath} -> {outputPath}\n");
        }
            else
            {
                Console.WriteLine($"Не удалось открыть файл {inputFilePath} после 6 попыток");
                if (logInFile)
                    File.AppendAllText(LogFilePath, $"Не удалось открыть файл {inputFilePath} после 6 попыток\n");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при конвертации {inputFilePath}: {ex.Message}");
            if (logInFile)
                File.AppendAllText(LogFilePath, $"Ошибка при конвертации {inputFilePath}: {ex.Message}\n");
            File.AppendAllText(ErrorLogFilePath, $"[{DateTime.UtcNow}] Ошибка при конвертации:\n{ex}\n");
            return false;
        }
        finally
        {
            if (workbook != null)
            {
                workbook.Close();
                Marshal.ReleaseComObject(workbook);
                workbook = null;
            }
    }
        return true;
    }

    private static Workbook? TryOpenWorkbook(string inputFilePath, Application excelApp, int maxRetries, bool verbose, bool logInFile)
    {
        int attempt = 0;
        Workbook? workbook = null;

        while (attempt < maxRetries)
        {
            try
            {
                workbook = excelApp.Workbooks.Open(inputFilePath);
                return workbook;
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x800AC472))
            {
                Console.WriteLine($"Ошибка при открытии файла {inputFilePath}: {ex.Message}\nПовторная попытка");
                if (logInFile)
                    File.AppendAllText(LogFilePath, $"Ошибка при открытии файла {inputFilePath}: {ex.Message}\nПовторная попытка\n");
                File.AppendAllText(ErrorLogFilePath, $"[{DateTime.UtcNow}] Ошибка при открытии файла:\n{ex}\nПовторная попытка\n");
        }

            attempt++;
            Thread.Sleep(1000 *  attempt);
        }

        return workbook;
    }
}