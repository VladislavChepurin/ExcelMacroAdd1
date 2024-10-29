using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
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
                MessageWarning("Функция работает только в \"Журнале учета НКУ\" текущего года. \n Пожайлуста откройте необходимую книгу Excel.",
                    "Имя книги не совпадает с целевой");
                return;
            }

            var firstRow = Cell.Row; // Вычисляем верхний элемент
            var countRow = Cell.Rows.Count; // Вычисляем кол-во выделенных строк
            var endRow = firstRow + countRow;
            do
            {
                //try
                //{
                    string sArticle = Convert.ToString(Worksheet.Cells[firstRow, 26].Value2);

                    if (!String.IsNullOrEmpty(sArticle))
                    {
                        var journalNku = await accessData.AccessJournalNku.GetEntityJournal(sArticle.ToLower());

                        if (journalNku is null)
                        {
                            Worksheet.Range["Z" + firstRow].Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
                            firstRow++;
                            continue;
                        }

                        Worksheet.Range["K" + firstRow].Value2 = journalNku.Ip.ToString();
                        Worksheet.Range["L" + firstRow].Value2 = journalNku.Climate ?? string.Empty;
                        Worksheet.Range["M" + firstRow].Value2 = journalNku.Reserve ?? string.Empty;
                        Worksheet.Range["N" + firstRow].Value2 = journalNku.Height ?? string.Empty;
                        Worksheet.Range["O" + firstRow].Value2 = journalNku.Width ?? string.Empty;
                        Worksheet.Range["P" + firstRow].Value2 = journalNku.Depth ?? string.Empty;
                        Worksheet.Range["AB" + firstRow].Value2 = journalNku.ExecutionBox.ExecutionValue ?? string.Empty;
                        Worksheet.Range["AD" + firstRow].Value2 = journalNku.MaterialBox.MaterialValue ?? string.Empty;
                    }
                //}

                //catch (DataException)
                //{
                //    MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                //        "Ошибка базы данных");
                //    return;
                //}
                //catch (Exception e)
                //{
                //    MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {e.Message}",
                //        "Ошибка базы данных");
                //    return;
                //}
                firstRow++;
            }
            while (endRow > firstRow);
        }
    }
}
