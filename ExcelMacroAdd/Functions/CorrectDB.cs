using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services;
using ExcelMacroAdd.UserException;
using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
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

        public override async Task StartAsync()
        {
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range selection = null;

            try
            {
                workbook = WorkBook;
                if (workbook?.Name != resources.NameFileJournal) // Проверка по имени книги
                {
                    MessageWarning(Properties.Resources.NotJornal, Properties.Resources.NameWorkbook);
                    return;
                }

                DialogResult dialogResult = MessageBox.Show(@"Вы уверены, что хотите изменить запись в БД? Пожалуйста будте очень внимательны, изменения коснуться всех пользователей.",
                                                            @"Контрольный вопрос", MessageBoxButtons.YesNo);
                if (dialogResult != DialogResult.Yes)
                {
                    return;
                }

                worksheet = Worksheet;
                selection = Cell;

                int currentRow = selection.Row; // Вычисляем верхний элемент
                string sCabinetArticle = ReadCellText(worksheet, currentRow, CabinetArticleColumn);

                if (string.IsNullOrEmpty(sCabinetArticle))
                {
                    MessageWarning("Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \n Артикул = ",
                        "Ошибка записи");
                    return;
                }

                if (!(await accessData.AccessJournalNku.GetEntityJournal(sCabinetArticle.ToLower()) is BoxBase journalNku))
                {
                    MessageWarning($"В базе данных такого артикула нет.\n Необходимо сначала его занести. \nАртикул = {sCabinetArticle}",
                        "Ошибка записи!");
                    return;
                }

                int.TryParse(ReadCellText(worksheet, currentRow, IPRatingColumn), out int sIPRating);
                string sClimaticCategory = ReadCellText(worksheet, currentRow, ClimaticCategoryColumn);
                string sMass = ReadCellText(worksheet, currentRow, MassColumn);
                string sEnclosureHeight = ReadCellText(worksheet, currentRow, EnclosureHeightColumn);
                string sEnclosureWidth = ReadCellText(worksheet, currentRow, EnclosureWidthColumn);
                string sEnclosureDepth = ReadCellText(worksheet, currentRow, EnclosureDepthColumn);
                string sCabinetMaterial = ReadCellText(worksheet, currentRow, CabinetMaterialTypeColumn);
                string sMountingType = ReadCellText(worksheet, currentRow, MountingTypeColumn);

                if (string.IsNullOrEmpty(sEnclosureHeight) || string.IsNullOrEmpty(sEnclosureWidth) || string.IsNullOrEmpty(sEnclosureDepth) || string.IsNullOrEmpty(sCabinetMaterial) || string.IsNullOrEmpty(sMountingType))
                {
                    MessageWarning($"Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \n Артикул = {sCabinetArticle}",
                        "Ошибка записи");
                    return;
                }

                var materialEntity = await accessData.AccessJournalNku.GetMaterialEntityByName(sCabinetMaterial)
                    ?? throw new DataBaseNotFoundValueException($"Введенный материал шкафа \"{sCabinetMaterial}\" недопустим, пожайлуста используйте значение \"Пластик\" или \"Металл\"");
                var executionEntity = await accessData.AccessJournalNku.GetExecutionEntityByName(sMountingType)
                    ?? throw new DataBaseNotFoundValueException($"Введенное исполнение шкафа \"{sMountingType}\" недопустимо, пожайлуста используйте значение \"напольное\", или \"навесное\", или \"встраиваемое\", или \"навесное для IT оборудования\", или \"напольное для IT оборудования\".");

                journalNku.Ip = sIPRating;
                journalNku.Climate = sClimaticCategory == "-" ? null : sClimaticCategory;
                journalNku.Weight = sMass == "-" ? null : sMass;
                journalNku.Height = sEnclosureHeight;
                journalNku.Width = sEnclosureWidth;
                journalNku.Depth = sEnclosureDepth;
                journalNku.Article = sCabinetArticle.ToLower();
                journalNku.MaterialBoxId = materialEntity.Id;
                journalNku.ExecutionBoxId = executionEntity.Id;

                await accessData.AccessJournalNku.WriteUpdateDb(journalNku);

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
            finally
            {
                ReleaseComObject(selection);
                ReleaseComObject(worksheet);
                ReleaseComObject(workbook);
            }
        }

        private static string ReadCellText(Worksheet worksheet, int row, int column)
        {
            Range cell = null;

            try
            {
                cell = (Range)worksheet.Cells[row, column];
                return Convert.ToString(cell.Value2);
            }
            finally
            {
                ReleaseComObject(cell);
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
