using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Interfaces;
using System;
using System.Data;
using System.Windows.Forms;

namespace ExcelMacroAdd.Functions
{
    internal class CorectDB : AbstractFunctions
    {
       // private JornalNKU jornalNKU;
        private readonly IJornalData jornalData;
        private readonly IResources resources;

        public CorectDB(IJornalData jornalData, IResources resources)
        {
            this.jornalData = jornalData;
            this.resources = resources;
        }

        public sealed override void Start()
        {
     
            if (application.ActiveWorkbook.Name != resources.NameFileJornal) // Проверка по имени книги
            {
                MessageWarning("Функция работает только в \"Журнале учета НКУ\" текущего года. \n Пожайлуста откройте необходимую книгу Excel.",
                     "Имя книги не совпадает с целевой");
                return;
            }
            DialogResult dialogResult = MessageBox.Show("Вы уверены, что хотите изменить запись в БД? \nИзменения коснуться всех пользователей.",
                                                        "Контрольный вопрос", MessageBoxButtons.YesNo);     
            if (dialogResult == DialogResult.Yes)
            {
                int firstRow;
                string sKlima, sReserve, sHeinght, sWidth, sDepth, sArticle, sExecution;

                firstRow = cell.Row;                 // Вычисляем верхний элемент
                sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);

                //using (DataContext db = new DataContext())
                //{
                //    var jornalNKUs = db.JornalNKU;
                try
                {
                    var jornalNKU = jornalData.GetEntityJornal(sArticle);
                    if (jornalNKU is null)
                    {
                        MessageWarning($"В базе данных такого артикула нет.\n Необходимо сначала его занести. \nАртикул = {sArticle}",
                        "Ошибка записи!");
                        return;
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
                        return;
                    }
                    jornalNKU.Ip = sIP;
                    jornalNKU.Klima = sKlima;
                    jornalNKU.Reserve = sReserve;
                    jornalNKU.Height = sHeinght;
                    jornalNKU.Width = sWidth;
                    jornalNKU.Depth = sDepth;
                    jornalNKU.Article = sArticle;
                    jornalNKU.Execution = sExecution;

                    jornalData.WriteUpdateDB((JornalNKU)jornalNKU);                                 

                    MessageInformation($"Запись успешно изменена! \nПоздравляем! \nАртикул = {sArticle}",
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
            }
        }
    }
}

