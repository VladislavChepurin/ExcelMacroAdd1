using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.Servises;
using System;
using System.Windows.Forms;

namespace ExcelMacroAdd.Functions
{
    internal class CorectDB : AbstractFunctions
    {
       
        private readonly IDBConect dBConect;

        public CorectDB(IDBConect dBConect)
        {
            this.dBConect = dBConect;
        }

        public override void Start()
        {
            DialogResult dialogResult = MessageBox.Show("Вы уверены, что хотите изменить запись в БД? \nИзменения коснуться всех пользователей.",
                                                        "Контрольный вопрос", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                int firstRow;
                string sIP, sKlima, sReserve, sHeinght, sWidth, sDepth, sArticle, sExecution;         
                try
                {
                    // Открываем соединение с базой данных    
                    dBConect?.OpenDB();

                    if (application.ActiveWorkbook.Name == dBConect?.ReadOnlyOneNoteDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2))            // Проверка по имени книги
                    {
                        firstRow = cell.Row;                 // Вычисляем верхний элемент
                        sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);

                        if (dBConect?.ReadOnlyOneNoteDB($"SELECT * FROM base WHERE Article = '{sArticle}';", 1) != null)
                        {
                            //Переписать! для этого есть структура DBTable
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
                                string queryUpdate = "SELECT * FROM base";
                                // Собираем запрос к БД   
                                string data = String.Format($"UPDATE base SET ip = '{sIP}', klima = '{sKlima}', reserve = '{sReserve}', height = '{sHeinght}'" +
                                    $", width = '{sWidth}', depth = '{sDepth}', execution = '{sExecution}' WHERE article = '{sArticle}';");
                                // Записываем в базу
                                dBConect?.UpdateNotesDB(queryUpdate, data);
                            }
                            else
                            {
                                MessageBox.Show(
                                "Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \n" +
                                "Артикул = " + sArticle,
                                "Ошибка записи",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                            }
                            // Закрываем соединение с базой данных
                            dBConect?.CloseDB();
                        }
                        else
                        {
                            MessageBox.Show(
                            "В базе данных такого артикула нет.\n Необходимо сначала его занести. \n" +
                            "Артикул = " + sArticle,
                            "Ошибка записи!",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning,
                            MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.DefaultDesktopOnly);
                        }
                    }
                }
                catch (Exception exception)
                {
                    MessageBox.Show(
                    exception.ToString(),
                    "Ошибка надстройки",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                }
            }
        }
    }
}
