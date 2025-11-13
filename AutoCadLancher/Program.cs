using AutoCadLancher;

namespace AutoCADBlockCopyConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("=== AutoCAD Block Copy Tool ===");
            Console.WriteLine();

            try
            {
                // Проверяем аргументы командной строки
                if (args.Length == 0)
                {
                    ShowUsage();
                    WaitForExit();
                    return;
                }

                // Парсим аргументы
                string? sourceDrawing = null;
                string? outputDrawing = null;
                List<string> blockNames = [];

                for (int i = 0; i < args.Length; i++)
                {
                    switch (args[i].ToLower())
                    {
                        case "-source":
                        case "-s":
                            if (i + 1 < args.Length) sourceDrawing = args[++i];
                            break;
                        case "-output":
                        case "-o":
                            if (i + 1 < args.Length) outputDrawing = args[++i];
                            break;
                        case "-blocks":
                        case "-b":
                            while (i + 1 < args.Length && !args[i + 1].StartsWith("-"))
                            {
                                blockNames.Add(args[++i]);
                            }
                            break;
                        case "-help":
                        case "-h":
                            ShowUsage();
                            WaitForExit();
                            return;
                    }
                }

                // Валидация параметров
                if (string.IsNullOrEmpty(sourceDrawing))
                {
                    Console.WriteLine("Ошибка: Не указан исходный файл чертежа");
                    ShowUsage();
                    WaitForExit();
                    return;
                }

                if (!File.Exists(sourceDrawing))
                {
                    Console.WriteLine($"Ошибка: Файл не найден: {sourceDrawing}");
                    WaitForExit();
                    return;
                }

                if (blockNames.Count == 0)
                {
                    Console.WriteLine("Ошибка: Не указаны имена блоков для копирования");
                    ShowUsage();
                    WaitForExit();
                    return;
                }

                // Генерируем имя выходного файла если не указано
                if (string.IsNullOrEmpty(outputDrawing))
                {
                    outputDrawing = GenerateOutputFileName(sourceDrawing);
                    Console.WriteLine($"Выходной файл не указан, будет создан: {outputDrawing}");
                }

                // Запускаем копирование блоков
                Console.WriteLine($"Исходный файл: {sourceDrawing}");
                Console.WriteLine($"Выходной файл: {outputDrawing}");
                Console.WriteLine($"Блоки для копирования: {string.Join(", ", blockNames)}");
                Console.WriteLine();

                BlockCopier copier = new BlockCopier();
                bool success = copier.CopyBlocksToNewDrawing(sourceDrawing, outputDrawing, blockNames.ToArray());

                if (success)
                {
                    Console.WriteLine($"\n✅ Операция завершена успешно! Создан файл: {outputDrawing}");
                }
                else
                {
                    Console.WriteLine("\n❌ Операция завершена с ошибками.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"💥 Критическая ошибка: {ex.Message}");
                Console.WriteLine($"Детали: {ex.StackTrace}");
            }

            WaitForExit();
        }

        static void ShowUsage()
        {
            Console.WriteLine("Использование:");
            Console.WriteLine("  AutoCADBlockCopyConsole.exe -source <файл.dwg> -blocks <блок1> <блок2> ... [-output <выходной_файл.dwg>]");
            Console.WriteLine();
            Console.WriteLine("Параметры:");
            Console.WriteLine("  -source, -s    Исходный файл чертежа (.dwg)");
            Console.WriteLine("  -blocks, -b    Список имен блоков для копирования");
            Console.WriteLine("  -output, -o    Выходной файл (опционально)");
            Console.WriteLine("  -help, -h      Показать эту справку");
            Console.WriteLine();
            Console.WriteLine("Примеры:");
            Console.WriteLine("  AutoCADBlockCopyConsole.exe -s \"C:\\drawings\\source.dwg\" -b \"Стол\" \"Стул\" \"Окно\"");
            Console.WriteLine("  AutoCADBlockCopyConsole.exe -source \"input.dwg\" -blocks Блок1 Блок2 -output \"output.dwg\"");
        }

        static string GenerateOutputFileName(string sourcePath)
        {
            string? directory = Path.GetDirectoryName(sourcePath);
            string fileName = Path.GetFileNameWithoutExtension(sourcePath);
            string extension = Path.GetExtension(sourcePath);

            return Path.Combine(directory!, $"{fileName}_blocks_{DateTime.Now:yyyyMMdd_HHmmss}{extension}");
        }

        static void WaitForExit()
        {
            Console.WriteLine("\nНажмите любую клавишу для выхода...");
            Console.ReadKey();
        }
    }
}