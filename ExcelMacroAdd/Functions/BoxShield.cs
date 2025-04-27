using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services;
using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal sealed class BoxShield : AbstractFunctions
    {
        private readonly IFillingOutThePassportSettings resources;
        private readonly IJournalData accessData;

        public BoxShield(IJournalData accessData, IFillingOutThePassportSettings resources)
        {
            this.accessData = accessData;
            this.resources = resources;
        }

        public override async void Start()
        {
            if (Application.ActiveWorkbook.Name != resources.NameFileJournal) // Проверка по имени книги
            {
                MessageWarning(Properties.Resources.NotJornal, Properties.Resources.NameWorkbook);
                return;
            }

            var firstRow = Cell.Row; // Вычисляем верхний элемент
            var countRow = Cell.Rows.Count; // Вычисляем кол-во выделенных строк
            var endRow = firstRow + countRow - 1;
            var currentRow = firstRow;

            do
            {
                try
                {
                    string sCabinetArticle = Convert.ToString(Worksheet.Cells[firstRow, CabinetArticleColumn].Value2);

                    if (!String.IsNullOrEmpty(sCabinetArticle))
                    {
                        var journalNku = await accessData.AccessJournalNku.GetEntityJournal(sCabinetArticle.ToLower());

                        if (journalNku is null)
                        {
                            Worksheet.Range["Z" + currentRow].Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
                            currentRow++;
                            continue;
                        }

                        Worksheet.Cells[currentRow, IPRatingColumn].Value2 = journalNku.Ip.ToString();
                        Worksheet.Cells[currentRow, ClimaticCategoryColumn].Value2 = journalNku.Climate ?? "-";
                        Worksheet.Cells[currentRow, MassColumn].Value2 = journalNku.Weight ?? "-";
                        Worksheet.Cells[currentRow, EnclosureHeightColumn].Value2 = journalNku.Height ?? string.Empty;
                        Worksheet.Cells[currentRow, EnclosureWidthColumn].Value2 = journalNku.Width ?? string.Empty;
                        Worksheet.Cells[currentRow, EnclosureDepthColumn].Value2 = journalNku.Depth ?? string.Empty;
                        Worksheet.Cells[currentRow, CabinetMaterialTypeColumn].Value2 = journalNku.MaterialBox.MaterialValue ?? string.Empty;
                        Worksheet.Cells[currentRow, MountingTypeColumn].Value2 = journalNku.ExecutionBox.ExecutionValue ?? string.Empty;
                    }
                }

                catch (DataException ex)
                {
                    MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                        "Ошибка базы данных");
                    Logger.LogException(ex);
                    return;
                }
                catch (Exception ex)
                {
                    MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {ex.Message}",
                        "Ошибка базы данных");
                    Logger.LogException(ex);
                    return;
                }
                currentRow++;
            }
            while (currentRow <= endRow);
        }
    }
}
