using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.BusinessLayer.Models;
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
        private readonly IJournalNkuWriteService journalNkuWriteService;
        private readonly IFillingOutThePassportSettings resources;

        public CorrectDb(IJournalNkuWriteService journalNkuWriteService, IFillingOutThePassportSettings resources)
        {
            this.journalNkuWriteService = journalNkuWriteService;
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
                if (workbook?.Name != resources.NameFileJournal)
                {
                    MessageWarning(Properties.Resources.NotJornal, Properties.Resources.NameWorkbook);
                    return;
                }

                DialogResult dialogResult = MessageBox.Show(
                    @"Вы уверены, что хотите изменить запись в БД? Пожалуйста будьте очень внимательны, изменения коснуться всех пользователей.",
                    @"Контрольный вопрос",
                    MessageBoxButtons.YesNo);

                if (dialogResult != DialogResult.Yes)
                {
                    return;
                }

                worksheet = Worksheet;
                selection = Cell;

                int currentRow = selection.Row;
                string article = ReadCellText(worksheet, currentRow, CabinetArticleColumn);

                if (string.IsNullOrEmpty(article))
                {
                    MessageWarning(
                        "Одно из обязательных полей не заполнено. Пожалуйста заполните все поля и еще раз повторите запись. \n Артикул = ",
                        "Ошибка записи");
                    return;
                }

                int.TryParse(ReadCellText(worksheet, currentRow, IPRatingColumn), out int ipRating);
                string climaticCategory = ReadCellText(worksheet, currentRow, ClimaticCategoryColumn);
                string mass = ReadCellText(worksheet, currentRow, MassColumn);
                string enclosureHeight = ReadCellText(worksheet, currentRow, EnclosureHeightColumn);
                string enclosureWidth = ReadCellText(worksheet, currentRow, EnclosureWidthColumn);
                string enclosureDepth = ReadCellText(worksheet, currentRow, EnclosureDepthColumn);
                string cabinetMaterial = ReadCellText(worksheet, currentRow, CabinetMaterialTypeColumn);
                string mountingType = ReadCellText(worksheet, currentRow, MountingTypeColumn);

                if (string.IsNullOrEmpty(enclosureHeight)
                    || string.IsNullOrEmpty(enclosureWidth)
                    || string.IsNullOrEmpty(enclosureDepth)
                    || string.IsNullOrEmpty(cabinetMaterial)
                    || string.IsNullOrEmpty(mountingType))
                {
                    MessageWarning(
                        $"Одно из обязательных полей не заполнено. Пожалуйста заполните все поля и еще раз повторите запись. \n Артикул = {article}",
                        "Ошибка записи");
                    return;
                }

                var request = new JournalNkuWriteRequest
                {
                    Ip = ipRating,
                    Climate = climaticCategory,
                    Weight = mass,
                    Height = enclosureHeight,
                    Width = enclosureWidth,
                    Depth = enclosureDepth,
                    Article = article,
                    MaterialName = cabinetMaterial,
                    ExecutionName = mountingType
                };

                var result = await journalNkuWriteService.UpdateBoxAsync(request);
                if (result.Status == JournalNkuWriteStatus.NotFound)
                {
                    MessageWarning(
                        $"В базе данных такого артикула нет.\n Необходимо сначала его занести. \nАртикул = {article}",
                        "Ошибка записи!");
                    return;
                }

                MessageInformation(
                    $"Запись успешно изменена! \nПоздравляем! \nАртикул = {article}",
                    "Запись успешна!");
            }
            catch (DataBaseNotFoundValueException ex)
            {
                MessageError(
                    $"Произошла ошибка, скорее всего неправильно было указано одно из значений. {ex.Message}",
                    "Ошибка базы данных");
                Logger.LogException(ex);
            }
            catch (DataException ex)
            {
                MessageError(
                    "Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                    "Ошибка базы данных");
                Logger.LogException(ex);
            }
            catch (Exception ex)
            {
                MessageError(
                    $"Произошла непредвиденная ошибка, пожалуйста сделайте скриншот ошибки, и передайте его разработчику.\n {ex.Message}",
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
