using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Interfaces;
using System;
using System.Data;

namespace ExcelMacroAdd.Functions
{
    internal class AddBoxDb : AbstractFunctions
    {
        private readonly IResources resources;
        private readonly IJournalData accessData;

        public AddBoxDb(IJournalData accessData, IResources resources)
        {
            this.accessData = accessData;
            this.resources = resources;
        }
        public sealed override async void Start()
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
                try
                {
                    string sArticle = Convert.ToString(Worksheet.Cells[firstRow, 26].Value2);

                    var journalNku = await accessData.AccessJournalNku.GetEntityJournal(sArticle.ToLower());

                    if (!(journalNku is null))
                    {
                        MessageWarning($"В базе данных уже есть такой артикул.\n Создавать новую запись не нужно. \nАртикул = {sArticle}",
                            "Ошибка записи!");
                        firstRow++;
                        continue;
                    }

                    int.TryParse(Convert.ToString(Worksheet.Cells[firstRow, 11].Value2), out int sIp);
                    string sClimate = Convert.ToString(Worksheet.Cells[firstRow, 12].Value2);
                    string sReserve = Convert.ToString(Worksheet.Cells[firstRow, 13].Value2);
                    string sHeight = Convert.ToString(Worksheet.Cells[firstRow, 14].Value2);
                    string sWidth = Convert.ToString(Worksheet.Cells[firstRow, 15].Value2);
                    string sDepth = Convert.ToString(Worksheet.Cells[firstRow, 16].Value2);
                    sArticle = Convert.ToString(Worksheet.Cells[firstRow, 26].Value2);
                    string sExecution = Convert.ToString(Worksheet.Cells[firstRow, 29].Value2);

                    if (sClimate == null || sReserve == null || sHeight == null || sWidth == null || sDepth == null || sArticle == null || sExecution == null)
                    {
                        MessageWarning($"Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \n Артикул = {sArticle}",
                            "Ошибка записи");
                        firstRow++;
                        continue;
                    }

                    JournalNku journal = new JournalNku()
                    {
                        Ip = sIp,
                        Climate = sClimate,
                        Reserve = sReserve,
                        Height = sHeight,
                        Width = sWidth,
                        Depth = sDepth,
                        Article = sArticle.ToLower(),
                        Execution = sExecution,
                        Vendor = "None"
                    };

                    accessData.AccessJournalNku.AddValueDb(journal);

                    MessageInformation($"Успешно записано в базу данных. Теперь доступна новая запись.\n Поздравляем! \nАртикул = {sArticle}",
                               "Запись успешна!");
                }
                catch (DataException)
                {
                    MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                        "Ошибка базы данных");
                    return;
                }
                catch (Exception e)
                {
                    MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {e.Message}",
                        "Ошибка базы данных");
                    return;
                }            
                firstRow++;
            }
            while (endRow > firstRow);
        }
    }
}
