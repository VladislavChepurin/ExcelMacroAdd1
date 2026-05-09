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
        private readonly IJournalData accessData;

        public BoxShield(IJournalData accessData, IFillingOutThePassportSettings resources)
        {
            this.accessData = accessData;
            this.resources = resources;
        }

        public override async Task StartAsync()
        {
            if (Application.ActiveWorkbook.Name != resources.NameFileJournal) // Проверка по имени книги
            {
                MessageWarning(Properties.Resources.NotJornal, Properties.Resources.NameWorkbook);
                return;
            }

            var firstRow = Cell.Row; // Вычисляем верхний элемент
            var countRow = Cell.Rows.Count; // Вычисляем кол-во выделенных строк
            var endRow = firstRow + countRow - 1;

            Excel.Range articleRange = null;
            Excel.Range enclosureRange = null;
            Excel.Range typeRange = null;

            try
            {
                articleRange = Worksheet.Range[
                    Worksheet.Cells[firstRow, CabinetArticleColumn],
                    Worksheet.Cells[endRow, CabinetArticleColumn]];

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

                var journalsByArticle = await accessData.AccessJournalNku.GetEntityJournalBatch(articlesToLoad);

                enclosureRange = Worksheet.Range[
                    Worksheet.Cells[firstRow, IPRatingColumn],
                    Worksheet.Cells[endRow, EnclosureDepthColumn]];
                typeRange = Worksheet.Range[
                    Worksheet.Cells[firstRow, CabinetMaterialTypeColumn],
                    Worksheet.Cells[endRow, MountingTypeColumn]];

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

                HighlightMissingRows(missingRows);
            }
            catch (DataException ex)
            {
                MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                    "Ошибка базы данных");
                Logger.LogException(ex);
            }
            catch (Exception ex)
            {
                MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {ex.Message}",
                    "Ошибка базы данных");
                Logger.LogException(ex);
            }
            finally
            {
                ReleaseComObject(typeRange);
                ReleaseComObject(enclosureRange);
                ReleaseComObject(articleRange);
            }
        }

        private void HighlightMissingRows(IReadOnlyList<int> missingRows)
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
                        highlightRange = Worksheet.Range[
                            Worksheet.Cells[segmentStart, MissingHighlightColumn],
                            Worksheet.Cells[previousRow, MissingHighlightColumn]];
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

        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null)
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
    }
}
