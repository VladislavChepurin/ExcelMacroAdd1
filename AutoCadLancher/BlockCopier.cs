using System.Runtime.InteropServices;

namespace AutoCadLancher
{
    public class BlockCopier
    {
        public bool CopyBlocksToNewDrawing(string sourceDrawingPath, string outputDrawingPath, string[] blockNames)
        {
            dynamic? acadApp = null;
            dynamic? sourceDoc = null;
            dynamic? newDoc = null;

            try
            {
                // Преобразуем пути в полные
                sourceDrawingPath = Path.GetFullPath(sourceDrawingPath);
                outputDrawingPath = Path.GetFullPath(outputDrawingPath);

                // Проверяем существование исходного файла
                if (!File.Exists(sourceDrawingPath))
                {
                    Console.WriteLine($"❌ Файл не найден: {sourceDrawingPath}");
                    Console.WriteLine($"📁 Текущая директория: {Environment.CurrentDirectory}");
                    return false;
                }

                // Получаем экземпляр AutoCAD
                acadApp = GetAutoCADInstance();
                if (acadApp == null)
                {
                    Console.WriteLine("❌ Не удалось получить доступ к AutoCAD");
                    return false;
                }

                // Открываем исходный чертеж
                Console.WriteLine($"📂 Открываем исходный чертеж: {sourceDrawingPath}");
                try
                {
                    sourceDoc = acadApp.Documents.Open(sourceDrawingPath, false, null);
                    Console.WriteLine("✅ Исходный чертеж успешно открыт");
                }
                catch (System.Exception openEx)
                {
                    Console.WriteLine($"❌ Ошибка при открытии файла: {openEx.Message}");
                    Console.WriteLine($"💡 Убедитесь, что файл не заблокирован и доступен для чтения");
                    return false;
                }

                // Создаем новый чертеж
                Console.WriteLine("🆕 Создаем новый чертеж...");
                try
                {
                    // Пробуем разные варианты вызова метода Add
                    newDoc = CreateNewDocument(acadApp);
                    if (newDoc == null)
                    {
                        Console.WriteLine("❌ Не удалось создать новый чертеж");
                        return false;
                    }
                    Console.WriteLine("✅ Новый чертеж создан");
                }
                catch (System.Exception createDocEx)
                {
                    Console.WriteLine($"❌ Ошибка при создании нового чертежа: {createDocEx.Message}");
                    return false;
                }

                // Копируем блоки
                Console.WriteLine("🔧 Копируем блоки...");
                int copiedBlocksCount = CopyBlocksBetweenDocuments(sourceDoc, newDoc, blockNames);

                if (copiedBlocksCount == 0)
                {
                    Console.WriteLine("⚠ Ни один блок не был скопирован");
                }

                // Сохраняем новый чертеж
                Console.WriteLine($"💾 Сохраняем новый чертеж: {outputDrawingPath}");

                // Создаем папку если не существует
                string? outputDir = Path.GetDirectoryName(outputDrawingPath);
                if (!Directory.Exists(outputDir) && !string.IsNullOrEmpty(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                    Console.WriteLine($"📁 Создана директория: {outputDir}");
                }

                try
                {
                    newDoc.SaveAs(outputDrawingPath);
                    Console.WriteLine("✅ Файл успешно сохранен");
                }
                catch (System.Exception saveEx)
                {
                    Console.WriteLine($"❌ Ошибка при сохранении файла: {saveEx.Message}");
                    return false;
                }

                // Закрываем документы
                Console.WriteLine("📪 Закрываем документы...");
                try
                {
                    sourceDoc.Close(false);
                    newDoc.Close(true);
                    Console.WriteLine("✅ Документы закрыты");
                }
                catch (System.Exception closeEx)
                {
                    Console.WriteLine($"⚠ Ошибка при закрытии документов: {closeEx.Message}");
                }

                Console.WriteLine($"✅ Успешно скопировано блоков: {copiedBlocksCount} из {blockNames.Length}");

                return copiedBlocksCount > 0;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"❌ Критическая ошибка: {ex.Message}");
                Console.WriteLine($"Детали: {ex.StackTrace}");
                return false;
            }
            finally
            {
                // Освобождаем COM объекты в правильном порядке
                SafeReleaseComObject(sourceDoc);
                SafeReleaseComObject(newDoc);
                SafeReleaseComObject(acadApp);
                Console.WriteLine("🏁 Работа завершена");
            }
        }

        private dynamic CreateNewDocument(dynamic acadApp)
        {
            // Пробуем разные варианты создания нового документа
            // так как сигнатуры методов могут отличаться в разных версиях AutoCAD

            try
            {
                // Вариант 1: Без параметров (самый простой)
                return acadApp.Documents.Add();
            }
            catch
            {
                try
                {
                    // Вариант 2: С пустой строкой в качестве шаблона
                    return acadApp.Documents.Add("");
                }
                catch
                {
                    try
                    {
                        // Вариант 3: С явным указанием null
                        return acadApp.Documents.Add(null);
                    }
                    catch
                    {
                        try
                        {
                            // Вариант 4: С указанием шаблона и других параметров
                            return acadApp.Documents.Add("acad.dwt");
                        }
                        catch (System.Exception ex)
                        {
                            Console.WriteLine($"❌ Все методы создания документа не сработали: {ex.Message}");
                            return null;
                        }
                    }
                }
            }
        }

        private dynamic GetAutoCADInstance()
        {
            Console.WriteLine("🔄 Подключение к AutoCAD...");

            // Сначала пробуем создать новый экземпляр (это может подключиться к запущенному)
            dynamic acadApp = CreateAutoCADInstance();
            if (acadApp != null)
            {
                try
                {
                    // Проверяем, подключились ли мы к запущенному экземпляру
                    string version = acadApp.Version;
                    int docCount = acadApp.Documents.Count;

                    if (docCount > 0)
                    {
                        Console.WriteLine($"✅ Подключено к запущенному AutoCAD (версия {version})");
                    }
                    else
                    {
                        Console.WriteLine($"✅ Создан новый экземпляр AutoCAD (версия {version})");
                        acadApp.Visible = true;
                    }
                    return acadApp;
                }
                catch
                {
                    // Если не удалось проверить, все равно возвращаем экземпляр
                    Console.WriteLine("✅ Подключение к AutoCAD выполнено");
                    return acadApp;
                }
            }

            Console.WriteLine("❌ Не удалось подключиться к AutoCAD");
            return null;
        }

        private dynamic CreateAutoCADInstance()
        {
            try
            {
                Type acadType = Type.GetTypeFromProgID("AutoCAD.Application");
                if (acadType == null)
                {
                    Console.WriteLine("❌ Тип AutoCAD.Application не найден");
                    return null;
                }

                return Activator.CreateInstance(acadType);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"❌ Ошибка создания экземпляра AutoCAD: {ex.Message}");
                return null;
            }
        }

        private void SafeReleaseComObject(object comObject)
        {
            if (comObject != null)
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch
                {
                    // Игнорируем ошибки освобождения
                }
            }
        }

        private int CopyBlocksBetweenDocuments(dynamic sourceDoc, dynamic targetDoc, string[] blockNames)
        {
            int copiedCount = 0;
            var blockNameSet = new HashSet<string>(blockNames, StringComparer.OrdinalIgnoreCase);
            var availableBlocks = new List<string>();

            try
            {
                // Получаем коллекции блоков из документов
                dynamic sourceBlocks = sourceDoc.Blocks;
                dynamic targetBlocks = targetDoc.Blocks;

                Console.WriteLine($"🔍 Поиск {blockNames.Length} блоков в исходном чертеже...");

                // Сначала соберем информацию о всех доступных блоках
                try
                {
                    int totalBlocks = sourceBlocks.Count;
                    Console.WriteLine($"📊 Всего блоков в чертеже: {totalBlocks}");

                    for (int i = 0; i < totalBlocks; i++)
                    {
                        try
                        {
                            dynamic block = sourceBlocks.Item(i);
                            string blockName = block.Name;

                            if (!IsStandardBlock(blockName))
                            {
                                availableBlocks.Add(blockName);
                            }
                        }
                        catch
                        {
                            // Пропускаем проблемные блоки
                        }
                    }

                    Console.WriteLine($"📋 Нестандартных блоков найдено: {availableBlocks.Count}");
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine($"⚠ Не удалось прочитать список блоков: {ex.Message}");
                }

                // Теперь копируем запрошенные блоки
                foreach (string blockName in blockNames)
                {
                    try
                    {
                        // Проверяем существует ли блок в исходном документе
                        if (!availableBlocks.Any(b => string.Equals(b, blockName, StringComparison.OrdinalIgnoreCase)))
                        {
                            Console.WriteLine($"❌ Блок '{blockName}' не найден в исходном чертеже");
                            continue;
                        }

                        // Проверяем, существует ли уже блок в целевом документе
                        bool blockExists = false;
                        try
                        {
                            dynamic existingBlock = targetBlocks.Item(blockName);
                            blockExists = true;
                            Console.WriteLine($"ℹ Блок '{blockName}' уже существует в целевом чертеже");
                        }
                        catch
                        {
                            blockExists = false;
                        }

                        if (!blockExists)
                        {
                            try
                            {
                                Console.WriteLine($"🔄 Копируем блок '{blockName}'...");
                                targetDoc.Import(blockName, sourceDoc.Name, false);
                                copiedCount++;
                                Console.WriteLine($"✅ Блок '{blockName}' успешно скопирован");
                            }
                            catch (System.Exception importEx)
                            {
                                Console.WriteLine($"❌ Ошибка при копировании блока '{blockName}': {importEx.Message}");
                            }
                        }
                        else
                        {
                            copiedCount++; // Считаем что блок уже существует
                        }
                    }
                    catch (System.Exception blockEx)
                    {
                        Console.WriteLine($"⚠ Ошибка при обработке блока '{blockName}': {blockEx.Message}");
                    }
                }

                // Покажем какие блоки доступны если не все были найдены
                if (copiedCount < blockNames.Length)
                {
                    var notFoundBlocks = blockNames.Where(bn =>
                        !availableBlocks.Any(ab => string.Equals(ab, bn, StringComparison.OrdinalIgnoreCase))).ToList();

                    if (notFoundBlocks.Count > 0)
                    {
                        Console.WriteLine($"⚠ Следующие блоки не найдены: {string.Join(", ", notFoundBlocks)}");
                    }

                    // Покажем первые 10 доступных блоков для справки
                    if (availableBlocks.Count > 0)
                    {
                        Console.WriteLine($"💡 Доступные блоки в чертеже (первые 10):");
                        foreach (string block in availableBlocks.Take(10))
                        {
                            Console.WriteLine($"   - {block}");
                        }
                        if (availableBlocks.Count > 10)
                        {
                            Console.WriteLine($"   ... и еще {availableBlocks.Count - 10} блоков");
                        }
                    }
                }

                return copiedCount;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"❌ Ошибка при копировании блоков: {ex.Message}");
                return copiedCount;
            }
        }

        private bool IsStandardBlock(string blockName)
        {
            if (string.IsNullOrEmpty(blockName))
                return true;

            string[] standardBlocks = {
        "*MODEL_SPACE", "*PAPER_SPACE", "*PAPER_SPACE0",
        "_ArchTick", "_Open30", "_Open90", "_Dot", "_DotSmall",
        "_DotBlank", "_Small", "_Closed", "_ClosedBlank", "_Oblique",
        "_Origin", "_Origin2", "_Circle"
        };

            return standardBlocks.Any(b => string.Equals(b, blockName, StringComparison.OrdinalIgnoreCase)) ||
                   blockName.StartsWith('$') ||
                   blockName.StartsWith('*') ||
                   blockName.StartsWith("A$");
        }
    }
}