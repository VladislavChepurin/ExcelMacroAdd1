using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal sealed class BoxShield : AbstractFunctions
    {
        private const int MissingHighlightColumn = CabinetArticleColumn - 1;

        private readonly IFillingOutThePassportSettings resources;
        private readonly IJournalNkuService journalNkuService;

        public BoxShield(IJournalNkuService journalNkuService, IFillingOutThePassportSettings resources)
        {
            this.journalNkuService = journalNkuService;
            this.resources = resources;
        }

        public override async Task StartAsync()
        {
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range selection = null;
            Excel.Range selectedRows = null;
            Excel.Range articleRange = null;
            Excel.Range enclosureRange = null;
            Excel.Range typeRange = null;

            try
            {
                workbook = WorkBook;
                if (workbook?.Name != resources.NameFileJournal) // Проверка по имени книги
                {
                    MessageWarning(Properties.Resources.NotJornal, Properties.Resources.NameWorkbook);
                    return;
                }

                worksheet = Worksheet;
                selection = Cell;
                selectedRows = selection.Rows;

                int firstRow = selection.Row; // Вычисляем верхний элемент
                int countRow = selectedRows.Count; // Вычисляем кол-во выделенных строк
                int endRow = firstRow + countRow - 1;

                articleRange = GetRange(worksheet, firstRow, CabinetArticleColumn, endRow, CabinetArticleColumn);
                var articleValues = ConvertRangeToMatrix(articleRange.Value2, countRow, 1);
                var articleKeysByRow = new string[countRow];
                var articlesToLoad = new List<string>(countRow);

                for (int rowIndex = 0; rowIndex < countRow; rowIndex++)
                {
                    string article = NormalizeArticle(Convert.ToString(articleValues[rowIndex, 0]));
                    if (string.IsNullOrEmpty(article))
                    {
                        continue;
                    }

                    articleKeysByRow[rowIndex] = article;
                    articlesToLoad.Add(article);
                }

                if (articlesToLoad.Count == 0)
                {
                    return;
                }

                var journalsByArticle = await journalNkuService.GetEntityJournalBatch(articlesToLoad);

                enclosureRange = GetRange(worksheet, firstRow, IPRatingColumn, endRow, EnclosureDepthColumn);
                typeRange = GetRange(worksheet, firstRow, CabinetMaterialTypeColumn, endRow, MountingTypeColumn);

                var enclosureValues = ConvertRangeToMatrix(
                    enclosureRange.Value2,
                    countRow,
                    EnclosureDepthColumn - IPRatingColumn + 1);
                var typeValues = ConvertRangeToMatrix(
                    typeRange.Value2,
                    countRow,
                    MountingTypeColumn - CabinetMaterialTypeColumn + 1);
                var missingRows = new List<int>();
                bool hasUpdates = false;

                for (int rowIndex = 0; rowIndex < countRow; rowIndex++)
                {
                    string article = articleKeysByRow[rowIndex];
                    if (string.IsNullOrEmpty(article))
                    {
                        continue;
                    }

                    if (!journalsByArticle.TryGetValue(article, out var journalNku) || journalNku is null)
                    {
                        missingRows.Add(firstRow + rowIndex);
                        continue;
                    }

                    enclosureValues[rowIndex, 0] = journalNku.Ip.ToString();
                    enclosureValues[rowIndex, 1] = journalNku.Climate ?? "-";
                    enclosureValues[rowIndex, 2] = journalNku.Weight ?? "-";
                    enclosureValues[rowIndex, 3] = journalNku.Height ?? string.Empty;
                    enclosureValues[rowIndex, 4] = journalNku.Width ?? string.Empty;
                    enclosureValues[rowIndex, 5] = journalNku.Depth ?? string.Empty;
                    typeValues[rowIndex, 0] = journalNku.MaterialBox?.MaterialValue ?? string.Empty;
                    typeValues[rowIndex, 1] = journalNku.ExecutionBox?.ExecutionValue ?? string.Empty;
                    hasUpdates = true;
                }

                if (hasUpdates)
                {
                    enclosureRange.Value2 = enclosureValues;
                    typeRange.Value2 = typeValues;
                }

                HighlightMissingRows(worksheet, missingRows);
            }
            catch (DataException ex)
            {
                MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                    "Ошибка базы данных");
                Logger.LogException(ex);
            }
            catch (Exception ex)
            {
                MessageError($"Произошла непредвиденная ошибка, пожалуйста сделайте скриншот ошибки, и передайте его разработчику.\n {ex.Message}",
                    "Ошибка базы данных");
                Logger.LogException(ex);
            }
            finally
            {
                ReleaseComObject(typeRange);
                ReleaseComObject(enclosureRange);
                ReleaseComObject(articleRange);
                ReleaseComObject(selectedRows);
                ReleaseComObject(selection);
                ReleaseComObject(worksheet);
                ReleaseComObject(workbook);
            }
        }

        private void HighlightMissingRows(Excel.Worksheet worksheet, IReadOnlyList<int> missingRows)
        {
            if (missingRows.Count == 0)
            {
                return;
            }

            int segmentStart = missingRows[0];
            int previousRow = missingRows[0];

            for (int index = 1; index <= missingRows.Count; index++)
            {
                bool isBoundary = index == missingRows.Count || missingRows[index] != previousRow + 1;
                if (isBoundary)
                {
                    Excel.Range highlightRange = null;
                    try
                    {
                        highlightRange = GetRange(worksheet, segmentStart, MissingHighlightColumn, previousRow, MissingHighlightColumn);
                        highlightRange.Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
                    }
                    finally
                    {
                        ReleaseComObject(highlightRange);
                    }
                }

                if (index < missingRows.Count)
                {
                    previousRow = missingRows[index];
                    if (isBoundary)
                    {
                        segmentStart = missingRows[index];
                    }
                }
            }
        }

        private static object[,] ConvertRangeToMatrix(object rangeValue, int rowCount, int columnCount)
        {
            var matrix = new object[rowCount, columnCount];

            if (rangeValue is object[,] values)
            {
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int column = 1; column <= columnCount; column++)
                    {
                        matrix[row - 1, column - 1] = values[row, column];
                    }
                }
            }
            else if (rowCount == 1 && columnCount == 1)
            {
                matrix[0, 0] = rangeValue;
            }

            return matrix;
        }

        private static string NormalizeArticle(string article)
        {
            return string.IsNullOrWhiteSpace(article)
                ? null
                : article.Trim().ToLowerInvariant();
        }

        private static Excel.Range GetRange(Excel.Worksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            Excel.Range startCell = null;
            Excel.Range endCell = null;

            try
            {
                startCell = (Excel.Range)worksheet.Cells[startRow, startColumn];
                endCell = (Excel.Range)worksheet.Cells[endRow, endColumn];
                return worksheet.Range[startCell, endCell];
            }
            finally
            {
                ReleaseComObject(endCell);
                ReleaseComObject(startCell);
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
    }
}
