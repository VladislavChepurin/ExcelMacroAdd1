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

namespace ExcelMacroAdd.Functions
{
    internal sealed class AddBoxDb : AbstractFunctions
    {
        private readonly IFillingOutThePassportSettings resources;
        private readonly IJournalNkuWriteService journalNkuWriteService;

        public AddBoxDb(IJournalNkuWriteService journalNkuWriteService, IFillingOutThePassportSettings resources)
        {
            this.journalNkuWriteService = journalNkuWriteService;
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
                if (workbook?.Name != resources.NameFileJournal)
                {
                    MessageWarning(Properties.Resources.NotJornal, Properties.Resources.NameWorkbook);
                    return;
                }

                worksheet = Worksheet;
                selection = Cell;
                selectedRows = selection.Rows;

                int firstRow = selection.Row;
                int countRow = selectedRows.Count;
                int endRow = firstRow + countRow - 1;
                int currentRow = firstRow;

                while (currentRow <= endRow)
                {
                    try
                    {
                        string article = ReadCellText(worksheet, currentRow, CabinetArticleColumn);

                        if (string.IsNullOrEmpty(article))
                        {
                            MessageWarning(
                                "Одно из обязательных полей не заполнено. Пожалуйста заполните все поля и еще раз повторите запись. \n Артикул = ",
                                "Ошибка записи");
                            continue;
                        }

                        int.TryParse(ReadCellText(worksheet, currentRow, IPRatingColumn), out int ip);
                        string climate = ReadCellText(worksheet, currentRow, ClimaticCategoryColumn);
                        string mass = ReadCellText(worksheet, currentRow, MassColumn);
                        string height = ReadCellText(worksheet, currentRow, EnclosureHeightColumn);
                        string width = ReadCellText(worksheet, currentRow, EnclosureWidthColumn);
                        string depth = ReadCellText(worksheet, currentRow, EnclosureDepthColumn);
                        string material = ReadCellText(worksheet, currentRow, CabinetMaterialTypeColumn);
                        string mountingType = ReadCellText(worksheet, currentRow, MountingTypeColumn);

                        if (string.IsNullOrEmpty(height)
                            || string.IsNullOrEmpty(width)
                            || string.IsNullOrEmpty(depth)
                            || string.IsNullOrEmpty(material)
                            || string.IsNullOrEmpty(mountingType))
                        {
                            MessageWarning(
                                $"Одно из обязательных полей не заполнено. Пожалуйста заполните все поля и еще раз повторите запись. \n Артикул = {article}",
                                "Ошибка записи");
                            continue;
                        }

                        var request = new JournalNkuWriteRequest
                        {
                            Ip = ip,
                            Climate = climate,
                            Weight = mass,
                            Height = height,
                            Width = width,
                            Depth = depth,
                            Article = article,
                            MaterialName = material,
                            ExecutionName = mountingType
                        };

                        var result = await journalNkuWriteService.AddBoxAsync(request);
                        if (result.Status == JournalNkuWriteStatus.AlreadyExists)
                        {
                            MessageWarning(
                                $"В базе данных уже есть такой артикул.\n Создавать новую запись не нужно. \nАртикул = {article}",
                                "Ошибка записи!");
                            continue;
                        }

                        MessageInformation(
                            $"Успешно записано в базу данных. Теперь доступна новая запись.\n Поздравляем! \nАртикул = {article}",
                            "Запись успешна!");
                    }
                    catch (DataBaseNotFoundValueException ex)
                    {
                        MessageError(
                            $"Произошла ошибка, скорее всего неправильно было указано исполнение шкафа. {ex.Message}",
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
