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

        public override async Task StartAsync()
        {
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range selection = null;
            Range selectedRows = null;

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
                int currentRow = firstRow;

                while (currentRow <= endRow)
                {
                    try
                    {
                        string sCabinetArticle = ReadCellText(worksheet, currentRow, CabinetArticleColumn);

                        if (string.IsNullOrEmpty(sCabinetArticle))
                        {
                            MessageWarning("Одно из обязательных полей не заполнено. Пожалуйста заполните все поля и еще раз повторите запись. \n Артикул = ",
                                "Ошибка записи");
                            continue;
                        }

                        var journalNku = await accessData.AccessJournalNku.GetEntityJournal(sCabinetArticle.ToLower());

                        if (!(journalNku is null))
                        {
                            MessageWarning($"В базе данных уже есть такой артикул.\n Создавать новую запись не нужно. \nАртикул = {sCabinetArticle}",
                                "Ошибка записи!");
                            continue;
                        }

                        int.TryParse(ReadCellText(worksheet, currentRow, IPRatingColumn), out int sIp);
                        string sClimate = ReadCellText(worksheet, currentRow, ClimaticCategoryColumn);
                        string sMass = ReadCellText(worksheet, currentRow, MassColumn);
                        string sHeight = ReadCellText(worksheet, currentRow, EnclosureHeightColumn);
                        string sWidth = ReadCellText(worksheet, currentRow, EnclosureWidthColumn);
                        string sDepth = ReadCellText(worksheet, currentRow, EnclosureDepthColumn);
                        string sMaterial = ReadCellText(worksheet, currentRow, CabinetMaterialTypeColumn);
                        string sMountingType = ReadCellText(worksheet, currentRow, MountingTypeColumn);

                        if (string.IsNullOrEmpty(sHeight) || string.IsNullOrEmpty(sWidth) || string.IsNullOrEmpty(sDepth) || string.IsNullOrEmpty(sMaterial))
                        {
                            MessageWarning($"Одно из обязательных полей не заполнено. Пожалуйста заполните все поля и еще раз повторите запись. \n Артикул = {sCabinetArticle}",
                                "Ошибка записи");
                            continue;
                        }

                        var materialEntity = await accessData.AccessJournalNku.GetMaterialEntityByName(sMaterial)
                            ?? throw new DataBaseNotFoundValueException($"Введенный материал шкафа \"{sMaterial}\" недопустим, пожалуйста используйте значение \"Пластик\", или  \"Металл\", или \"Композит\"");
                        var executionEntity = await accessData.AccessJournalNku.GetExecutionEntityByName(sMountingType)
                            ?? throw new DataBaseNotFoundValueException($"Введенное исполнение шкафа \"{sMountingType}\" недопустимо, пожалуйста используйте значение \"напольное\", или \"навесное\", или \"встраиваемое\", или \"навесное для IT оборудования\", или \"напольное для IT оборудования\".");

                        BoxBase journal = new BoxBase
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

                        await accessData.AccessJournalNku.AddValueDb(journal);

                        MessageInformation($"Успешно записано в базу данных. Теперь доступна новая запись.\n Поздравляем! \nАртикул = {sCabinetArticle}",
                            "Запись успешна!");
                    }
                    catch (DataBaseNotFoundValueException ex)
                    {
                        MessageError($"Произошла ошибка, скорее всего неправильно было указано исполнение шкафа. {ex.Message}",
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
                        MessageError($"Произошла непредвиденная ошибка, пожалуйста сделайте скриншот ошибки, и передайте его разработчику.\n {ex.Message}",
                            "Ошибка базы данных");
                        Logger.LogException(ex);
                    }
                    finally
                    {
                        currentRow++;
                    }
                }
            }
            finally
            {
                ReleaseComObject(selectedRows);
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