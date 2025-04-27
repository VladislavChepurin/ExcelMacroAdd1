using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services;
using ExcelMacroAdd.UserException;
using System;
using System.Data;
using System.Windows.Forms;

namespace ExcelMacroAdd.Functions
{
    internal sealed class CorrectDb : AbstractFunctions
    {
        private readonly IJournalData accessData;
        private readonly IFillingOutThePassportSettings resources;

        public CorrectDb(IJournalData accessData, IFillingOutThePassportSettings resources)
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
            DialogResult dialogResult = MessageBox.Show(@"Вы уверены, что хотите изменить запись в БД? Пожалуйста будте очень внимательны, изменения коснуться всех пользователей.",
                                                        @"Контрольный вопрос", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                var currentRow = Cell.Row; // Вычисляем верхний элемент
                string sCabinetArticle = Convert.ToString(Worksheet.Cells[currentRow, CabinetArticleColumn].Value2);

                try
                {
                    if (!(await accessData.AccessJournalNku.GetEntityJournal(sCabinetArticle.ToLower()) is BoxBase journalNku))
                    {
                        MessageWarning($"В базе данных такого артикула нет.\n Необходимо сначала его занести. \nАртикул = {sCabinetArticle}",
                        "Ошибка записи!");
                        return;
                    }

                    int.TryParse(Convert.ToString(Worksheet.Cells[currentRow, IPRatingColumn].Value2), out int sIPRating);
                    string sClimaticCategory = Convert.ToString(Worksheet.Cells[currentRow, ClimaticCategoryColumn].Value2);
                    string sMass = Convert.ToString(Worksheet.Cells[currentRow, MassColumn].Value2);
                    string sEnclosureHeight = Convert.ToString(Worksheet.Cells[currentRow, EnclosureHeightColumn].Value2);
                    string sEnclosureWidth = Convert.ToString(Worksheet.Cells[currentRow, EnclosureWidthColumn].Value2);
                    string sEnclosureDepth = Convert.ToString(Worksheet.Cells[currentRow, EnclosureDepthColumn].Value2);                   
                    string sCabinetMaterial = Convert.ToString(Worksheet.Cells[currentRow, CabinetMaterialTypeColumn].Value2);
                    string sMountingType = Convert.ToString(Worksheet.Cells[currentRow, MountingTypeColumn].Value2);

                    if (string.IsNullOrEmpty(sEnclosureHeight) || string.IsNullOrEmpty(sEnclosureWidth) || string.IsNullOrEmpty(sEnclosureDepth) || string.IsNullOrEmpty(sCabinetArticle) || string.IsNullOrEmpty(sCabinetMaterial) || string.IsNullOrEmpty(sMountingType))
                    {
                        MessageWarning($"Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \n Артикул = {sCabinetArticle}",
                            "Ошибка записи");
                        return;
                    }

                    var materialEntity = await accessData.AccessJournalNku.GetMaterialEntityByName(sCabinetMaterial) ?? throw new DataBaseNotFoundValueException($"Введенный материал шкафа \"{sCabinetMaterial}\" недопустим, пожайлуста используйте значение \"Пластик\" или \"Металл\"");
                    var executionEntity = await accessData.AccessJournalNku.GetExecutionEntityByName(sMountingType) ?? throw new DataBaseNotFoundValueException($"Введенное исполнение шкафа \"{sMountingType}\" недопустимо, пожайлуста используйте значение \"напольное\", или \"навесное\", или \"встраиваемое\", или \"навесное для IT оборудования\", или \"напольное для IT оборудования\".");

                    journalNku.Ip = sIPRating;
                    journalNku.Climate = sClimaticCategory == "-" ? null : sClimaticCategory;
                    journalNku.Weight = sMass == "-" ? null : sMass;
                    journalNku.Height = sEnclosureHeight;
                    journalNku.Width = sEnclosureWidth;
                    journalNku.Depth = sEnclosureDepth;
                    journalNku.Article = sCabinetArticle.ToLower();
                    journalNku.MaterialBoxId = materialEntity.Id;
                    journalNku.ExecutionBoxId = executionEntity.Id;

                    accessData.AccessJournalNku.WriteUpdateDb(journalNku);

                    MessageInformation($"Запись успешно изменена! \nПоздравляем! \nАртикул = {sCabinetArticle}",
                                "Запись успешна!");
                }

                catch (DataBaseNotFoundValueException ex)
                {
                    MessageError($"Произошла ошибка, скорее всего непавильно было указано одно из значений. {ex.Message}",
                        "Ошибка базы данных");
                    Logger.LogException(ex);
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
            }
        }
    }
}
