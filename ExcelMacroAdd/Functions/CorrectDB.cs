using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Interfaces;
using System;
using System.Data;
using System.Windows.Forms;

namespace ExcelMacroAdd.Functions
{
    internal class CorrectDb : AbstractFunctions
    {       
        private readonly IJournalData journalData;
        private readonly IResources resources;

        public CorrectDb(IJournalData journalData, IResources resources)
        {
            this.journalData = journalData;
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
            DialogResult dialogResult = MessageBox.Show(@"Вы уверены, что хотите изменить запись в БД? Пожалуйста будте очень внимательны, изменения коснуться всех пользователей.",
                                                        @"Контрольный вопрос", MessageBoxButtons.YesNo);     
            if (dialogResult == DialogResult.Yes)
            {
                var firstRow = Cell.Row; // Вычисляем верхний элемент
                string sArticle = Convert.ToString(Worksheet.Cells[firstRow, 26].Value2);

                try
                {
                    var jornalNku = await journalData.GetEntityJournal(sArticle);
                    if (jornalNku is null)
                    {
                        MessageWarning($"В базе данных такого артикула нет.\n Необходимо сначала его занести. \nАртикул = {sArticle}",
                        "Ошибка записи!");
                        return;
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
                        return;
                    }
                    jornalNku.Ip = sIp;
                    jornalNku.Climate = sClimate;
                    jornalNku.Reserve = sReserve;
                    jornalNku.Height = sHeight;
                    jornalNku.Width = sWidth;
                    jornalNku.Depth = sDepth;
                    jornalNku.Article = sArticle;
                    jornalNku.Execution = sExecution;

                    journalData.WriteUpdateDB((JournalNku)jornalNku);                                 

                    MessageInformation($"Запись успешно изменена! \nПоздравляем! \nАртикул = {sArticle}",
                                "Запись успешна!");
                }
                catch (DataException)
                {
                    MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                        "Ошибка базы данных");
                }
                catch (Exception e)
                {
                    MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {e.Message}",
                        "Ошибка базы данных");
                }
            }
        }
    }
}

