using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Interfaces;
using System;
using System.Data;

namespace ExcelMacroAdd.Functions
{
    internal class AddBoxDB : AbstractFunctions
    {
        private readonly IResources resources;
        private readonly IJornalData jornalData;

        public AddBoxDB(IJornalData jornalData, IResources resources)
        {
            this.jornalData = jornalData;
            this.resources = resources;
        }
        public sealed override async void Start()
        {
            if (application.ActiveWorkbook.Name != resources.NameFileJornal) // Проверка по имени книги
            {
                MessageWarning("Функция работает только в \"Журнале учета НКУ\" текущего года. \n Пожайлуста откройте необходимую книгу Excel.",
                    "Имя книги не совпадает с целевой");
                return;
            }
            int firstRow, countRow, endRow;
            string sKlima, sReserve, sHeinght, sWidth, sDepth, sArticle, sExecution;

            firstRow = cell.Row;                 // Вычисляем верхний элемент
            countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
            endRow = firstRow + countRow;

            do
            {
                try
                {
                    sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);

                    var jornalNKU = await jornalData.GetEntityJornal(sArticle);

                    if (!(jornalNKU is null))
                    {
                        MessageWarning($"В базе данных уже есть такой артикул.\n Создавать новую запись не нужно. \nАртикул = {sArticle}",
                            "Ошибка записи!");
                        firstRow++;
                        continue;
                    }

                    int.TryParse(Convert.ToString(worksheet.Cells[firstRow, 11].Value2), out int sIP);
                    sKlima = Convert.ToString(worksheet.Cells[firstRow, 12].Value2);
                    sReserve = Convert.ToString(worksheet.Cells[firstRow, 13].Value2);
                    sHeinght = Convert.ToString(worksheet.Cells[firstRow, 14].Value2);
                    sWidth = Convert.ToString(worksheet.Cells[firstRow, 15].Value2);
                    sDepth = Convert.ToString(worksheet.Cells[firstRow, 16].Value2);
                    sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);
                    sExecution = Convert.ToString(worksheet.Cells[firstRow, 29].Value2);

                    if (sKlima == null || sReserve == null || sHeinght == null || sWidth == null || sDepth == null || sArticle == null || sExecution == null)
                    {
                        MessageWarning($"Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \n Артикул = {sArticle}",
                            "Ошибка записи");
                        firstRow++;
                        continue;
                    }

                    JornalNKU jornal = new JornalNKU()
                    {
                        Ip = sIP,
                        Klima = sKlima,
                        Reserve = sReserve,
                        Height = sHeinght,
                        Width = sWidth,
                        Depth = sDepth,
                        Article = sArticle,
                        Execution = sExecution,
                        Vendor = "None"
                    };

                    jornalData.AddValueDB(jornal);

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
