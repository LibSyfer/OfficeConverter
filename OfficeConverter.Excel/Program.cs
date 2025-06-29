﻿using CommandLine;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Polly;
using Polly.Retry;
using System.Diagnostics;
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

    private static int RetryCount = 10;

    private static RetryPolicy LicenceRetryPolicy = Policy.Handle<COMException>().Retry();

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
        Application? excelApp = null;
        int exitCode = 0;

        LicenceRetryPolicy = Policy.Handle<COMException>()
        .WaitAndRetry(
            retryCount: RetryCount,
            sleepDurationProvider: attempt => TimeSpan.FromSeconds(2 * attempt),
            onRetry: (exception, delay, retryCount, context) =>
            {
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine($"Работа excel блокируется ({context.OperationKey}). Попытка восстановить: {retryCount}");
                Console.WriteLine("Закройте все всплывающие окна excel, блокирующие работу");
                Console.ResetColor();
                if (options.LogInFile)
                    File.AppendAllText(LogFilePath, $"Работа excel блокируется ({context.OperationKey}). Попытка восстановить: {retryCount}\n");
                File.AppendAllText(ErrorLogFilePath, $"Работа excel блокируется ({context.OperationKey}): {exception}\nПопытка восстановить: {retryCount}\n");
            });

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
                File.AppendAllText(LogFilePath, $"Начало конвертации {DateTime.UtcNow}\n");
                File.AppendAllText(ErrorLogFilePath, $"Начало конвертации {DateTime.UtcNow}\n");

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

            excelApp = new Application();
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
        catch (Exception ex)
        {
            Console.WriteLine($"Глобальная ошибка: {ex.Message}");
            if (options.LogInFile)
            {
                File.AppendAllText(LogFilePath, $"Глобальная ошибка: {ex.Message}\n");
                File.AppendAllText(ErrorLogFilePath, $"Глобальная ошибка:\n{ex}\n");
            }
            exitCode = 1;
        }
        finally
        {
            if (options.Verbose)
                Console.WriteLine("Очистка COM объекта фонового приложения excel");
            if (options.LogInFile)
                File.AppendAllText(LogFilePath, "Очистка COM объекта фонового приложения excel\n");
            try { excelApp?.Quit(); } catch { }
            Marshal.FinalReleaseComObject(excelApp);
            excelApp = null;

            if (options.Verbose)
                Console.WriteLine("Очистка процессов excel");
            if (options.LogInFile)
                File.AppendAllText(LogFilePath, "Очистка процессов excel\n");
            KillExcelProcesses(options.Verbose, options.LogInFile);
        }
        Environment.Exit(exitCode);
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

            LicenceRetryPolicy.Execute((context) => workbook = excelApp.Workbooks.Open(inputFilePath), new Context("excel.workbooks.Open"));
            LicenceRetryPolicy.Execute((context) =>
                    workbook!.SaveAs(
                        Filename: outputPath,
                        FileFormat: XlFileFormat.xlOpenXMLWorkbook,
                        ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges,
                        Local: true,
                        AddToMru: false),
                    new Context("workbook.SaveAs")
                    );

            Console.WriteLine($"Конвертирован: {inputFilePath} -> {outputPath}");
            if (logInFile)
                File.AppendAllText(LogFilePath, $"Конвертирован: {inputFilePath} -> {outputPath}\n");
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800AC472))
        {
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine($"Критическая ошибка: не удалось снять блокировку excel спустя {RetryCount} попыток");
            Console.ResetColor();
            if (logInFile)
            {
                File.AppendAllText(LogFilePath, $"Критическая ошибка: не удалось снять блокировку excel спустя {RetryCount} попыток: {ex.Message}\n");
                File.AppendAllText(ErrorLogFilePath, $"Критическая ошибка: не удалось снять блокировку excel спустя {RetryCount} попыток:\n{ex}\n");
            }
            throw;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при конвертации {inputFilePath}: {ex.Message}");
            if (logInFile)
            {
                File.AppendAllText(LogFilePath, $"Ошибка при конвертации {inputFilePath}: {ex.Message}\n");
                File.AppendAllText(ErrorLogFilePath, $"Ошибка при конвертации {inputFilePath}:\n{ex}\n");
            }
            return false;
        }
        finally
        {
            if (workbook != null)
            {
                if (logInFile)
                    File.AppendAllText(LogFilePath, $"Очистка COM объектов файла {inputFilePath}\n");
                LicenceRetryPolicy.Execute((context) => workbook.Close(), new Context("workbook.Close"));
                Marshal.FinalReleaseComObject(workbook);
                workbook = null;
            }
        }
        return true;
    }

    private static void KillExcelProcesses(bool verbose, bool logInFile)
    {
        int processCount = 0;
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            try
            {
                if (process.StartTime > Process.GetCurrentProcess().StartTime)
                {
                    process.Kill();
                }
                processCount++;
            }
            catch (Exception ex)
            {
                if (logInFile && verbose)
                {
                    File.AppendAllText(LogFilePath, $"Не удалось остановить процесс:\n{ex}\n");
                    File.AppendAllText(ErrorLogFilePath, $"Не удалось остановить процесс:\n{ex}\n");
                }
            }
        }

        if (verbose)
            Console.WriteLine($"{processCount} процессов остановлено");
        if (logInFile)
            File.AppendAllText(LogFilePath, $"{processCount} процессов остановлено\n");
    }
}