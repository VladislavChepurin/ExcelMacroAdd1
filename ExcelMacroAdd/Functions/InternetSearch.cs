using ExcelMacroAdd.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelMacroAdd.Functions
{
    internal sealed class InternetSearch : AbstractFunctions
    {
        private readonly string searchLink;
        public InternetSearch(string searchLink)
        {
            this.searchLink = searchLink;
        }

        public override void Start()
        {
            // Получаем значение ячейки/диапазона
            var cellValue = Cell.Value;

            if (cellValue != null)
            {
                string searchQuery;

                // Обрабатываем случай выделенного диапазона ячеек
                if (cellValue is object[,] valueArray)
                {
                    var queryParts = new List<string>();

                    // Проходим по всем элементам массива
                    foreach (var item in valueArray)
                    {
                        // Безопасное преобразование в строку
                        if (item != null)
                        {
                            var stringValue = item.ToString();

                            // Пропускаем пустые строки
                            if (!string.IsNullOrWhiteSpace(stringValue))
                            {
                                queryParts.Add(stringValue);
                            }
                        }
                    }

                    searchQuery = string.Join(" ", queryParts);
                }
                else // Одиночная ячейка
                {
                    searchQuery = cellValue.ToString();
                }

                if (!string.IsNullOrWhiteSpace(searchQuery))
                {
                    try
                    {
                        // Кодируем запрос для URL
                        string encodedQuery = Uri.EscapeDataString(searchQuery);
                        string searchUrl = String.Concat(searchLink, encodedQuery);

                        // Запускаем процесс с обработкой ошибок
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = searchUrl,
                            UseShellExecute = true // Важно для работы в современных версиях .NET
                        });
                    }
                    catch (Exception ex)
                    {
                        // Обработка ошибок (логирование, уведомление пользователя)                    
                        MessageError($"Ошибка при открытии браузера: {ex.Message}", "Ошибка");
                        Logger.LogException(ex);
                    }
                }
            }
        }
    }
}
