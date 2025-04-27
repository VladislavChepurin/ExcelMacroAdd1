using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services;
using ExcelMacroAdd.UserException;
using Microsoft.Office.Interop.Word;
using System;
using System.Data;

namespace ExcelMacroAdd.Functions
{
    internal sealed class AddBoxDb : AbstractFunctions
    {
        private readonly IFillingOutThePassportSettings resources;
        private readonly IJournalData accessData;

        public AddBoxDb(IJournalData accessData, IFillingOutThePassportSettings resources)
        {
            this.accessData = accessData;
            this.resources = resources;
        }
        public override async void Start()
        {
            if (Application.ActiveWorkbook?.Name != resources.NameFileJournal) // Проверка по имени книги
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
                    string sCabinetArticle = Convert.ToString(Worksheet.Cells[currentRow, CabinetArticleColumn].Value2);
                    var journalNku = await accessData.AccessJournalNku.GetEntityJournal(sCabinetArticle.ToLower());

                    if (!(journalNku is null))
                    {
                        MessageWarning($"В базе данных уже есть такой артикул.\n Создавать новую запись не нужно. \nАртикул = {sCabinetArticle}",
                            "Ошибка записи!");
                        currentRow++;
                        continue;
                    }

                    int.TryParse(Convert.ToString(Worksheet.Cells[currentRow, IPRatingColumn].Value2), out int sIp);
                    string sClimate = Convert.ToString(Worksheet.Cells[currentRow, ClimaticCategoryColumn].Value2);
                    string sMass = Convert.ToString(Worksheet.Cells[currentRow, MassColumn].Value2);
                    string sHeight = Convert.ToString(Worksheet.Cells[currentRow, EnclosureHeightColumn].Value2);
                    string sWidth = Convert.ToString(Worksheet.Cells[currentRow, EnclosureWidthColumn].Value2);
                    string sDepth = Convert.ToString(Worksheet.Cells[currentRow, EnclosureDepthColumn].Value2);
                    sCabinetArticle = Convert.ToString(Worksheet.Cells[currentRow, CabinetArticleColumn].Value2);
                    string sMaterial = Convert.ToString(Worksheet.Cells[currentRow, CabinetMaterialTypeColumn].Value2);
                    string sMountingType = Convert.ToString(Worksheet.Cells[currentRow, MountingTypeColumn].Value2);

                    if (string.IsNullOrEmpty(sHeight) || string.IsNullOrEmpty(sWidth) || string.IsNullOrEmpty(sDepth) || string.IsNullOrEmpty(sCabinetArticle) || string.IsNullOrEmpty(sMaterial))
                    {
                        MessageWarning($"Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \n Артикул = {sCabinetArticle}",
                            "Ошибка записи");
                        currentRow++;
                        continue;
                    }

                    var materialEntity = await accessData.AccessJournalNku.GetMaterialEntityByName(sMaterial) ?? throw new DataBaseNotFoundValueException($"Введенный материал шкафа \"{sMaterial}\" недопустим, пожайлуста используйте значение \"Пластик\", или  \"Металл\", или \"Композит\"");
                    var executionEntity = await accessData.AccessJournalNku.GetExecutionEntityByName(sMountingType) ?? throw new DataBaseNotFoundValueException($"Введенное исполнение шкафа \"{sMountingType}\" недопустимо, пожайлуста используйте значение \"напольное\", или \"навесное\", или \"встраиваемое\", или \"навесное для IT оборудования\", или \"напольное для IT оборудования\".");

                    BoxBase journal = new BoxBase()
                    {
                        Ip = sIp,
                        Climate = sClimate == "-" ? null : sClimate,
                        Weight = sMass == "-" ? null : sMass,
                        Height = sHeight,
                        Width = sWidth,
                        Depth = sDepth,
                        Article = sCabinetArticle.ToLower(),
                        MaterialBoxId = materialEntity.Id,
                        ExecutionBoxId = executionEntity.Id
                    };

                    accessData.AccessJournalNku.AddValueDb(journal);

                    MessageInformation($"Успешно записано в базу данных. Теперь доступна новая запись.\n Поздравляем! \nАртикул = {sCabinetArticle}",
                               "Запись успешна!");
                }

                catch (DataBaseNotFoundValueException ex)
                {
                    MessageError($"Произошла ошибка, скорее всего непавильно было указано исполнение шкафа. {ex.Message}",
                        "Ошибка базы данных");
                    Logger.LogException(ex);
                    continue;
                }

                catch (DataException ex)
                {
                    MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                        "Ошибка базы данных");
                    Logger.LogException(ex);
                    continue;
                }

                catch (Exception ex)
                {
                    MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {ex.Message}",
                        "Ошибка базы данных");
                    Logger.LogException(ex);
                    continue;
                }
                currentRow++;
            }
            while (currentRow <= endRow);
        }
    }
}
