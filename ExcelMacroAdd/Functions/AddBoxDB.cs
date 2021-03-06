using ExcelMacroAdd.Interfaces;
using System;

namespace ExcelMacroAdd.Functions
{
    internal class AddBoxDB : AbstractFunctions
    {
        private readonly IDBConect dBConect;

        public AddBoxDB(IDBConect dBConect)
        {
            this.dBConect = dBConect;
        }
        protected internal override void Start()
        {
            dBConect?.OpenDB();
            if (application.ActiveWorkbook.Name != dBConect?.ReadOnlyOneNoteDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2)) // Проверка по имени книги
            {
                MessageWarning("Функция работает только в \"Журнале учета НКУ\" текущего года. \n Пожайлуста откройте необходимую книгу Excel.",
                    "Имя книги не совпадает с целевой");
                dBConect?.CloseDB();
                return;
            }

            int firstRow, countRow, endRow;
            string sIP, sKlima, sReserve, sHeinght, sWidth, sDepth, sArticle, sExecution;
            try
            {
                firstRow = cell.Row;                 // Вычисляем верхний элемент
                countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
                endRow = firstRow + countRow;
                do
                {
                    //Переписать! Для этого есть структура DBTable
                    sIP = Convert.ToString(worksheet.Cells[firstRow, 11].Value2);
                    sKlima = Convert.ToString(worksheet.Cells[firstRow, 12].Value2);
                    sReserve = Convert.ToString(worksheet.Cells[firstRow, 13].Value2);
                    sHeinght = Convert.ToString(worksheet.Cells[firstRow, 14].Value2);
                    sWidth = Convert.ToString(worksheet.Cells[firstRow, 15].Value2);
                    sDepth = Convert.ToString(worksheet.Cells[firstRow, 16].Value2);
                    sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);
                    sExecution = Convert.ToString(worksheet.Cells[firstRow, 29].Value2);

                    // Если хоть одно поле не заполнено, то записи в базу нет
                    if (sIP != null && sKlima != null && sReserve != null && sHeinght != null
                        && sWidth != null && sDepth != null && sArticle != null && sExecution != null)
                    {
                        if (dBConect.ReadOnlyOneNoteDB($"SELECT * FROM base WHERE Article = '{sArticle}';", 1) is null)
                        {
                            //Сборка запроса к БД
                            string commandText = $"INSERT INTO base (ip, klima, reserve, height, width, depth, article, execution, vendor)" +
                                  $" VALUES ('{sIP}', '{sKlima}', '{sReserve}', '{sHeinght}', '{sWidth}', '{sDepth}', '{sArticle}', '{sExecution}', 'None');";
                            //Оправка запроса к БД
                            dBConect.UpdateNotesDB("SELECT * FROM base", commandText);
                            worksheet.get_Range("Z" + firstRow).Interior.ColorIndex = 0;
                            MessageInformation($"Успешно записано в базу данных. Теперь доступна новая запись.\n Поздравляем! \nАртикул = {sArticle}",
                                "Запись успешна!");
                        }
                        else
                        { 
                            MessageWarning($"В базе данных уже есть такой артикул.\n Создавать новую запись не нужно. \nАртикул = {sArticle}",
                                "Ошибка записи!");
                        }
                    }
                    else
                    {
                        MessageWarning($"Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \nАртикул = {sArticle}",
                            "Ошибка записи!");
                    }
                    firstRow++;
                }
                while (endRow > firstRow);
                dBConect.CloseDB();
            }
            catch (Exception exception)
            {
                MessageError(exception.ToString(),
                    "Ошибка надсройки");
            }
        }
    }
}
