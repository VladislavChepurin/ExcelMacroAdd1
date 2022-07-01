using ExcelMacroAdd.Servises;
using System;
using System.Windows.Forms;

namespace ExcelMacroAdd.Functions
{
    internal class AddBoxDB : AbstractFunctions
    {
        private readonly Lazy<DBConect> dBConect;

        public AddBoxDB(Lazy<DBConect> dBConect)
        {
            this.dBConect = dBConect;
        }

        public override void Start()
        {
            int firstRow, countRow, endRow;
            string sIP, sKlima, sReserve, sHeinght, sWidth, sDepth, sArticle, sExecution;           
            try
            {
                // Открываем соединение с базой данных    
                dBConect.Value.OpenDB();
                // Проверка по имени книги
                if (application.ActiveWorkbook.Name == dBConect.Value.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2))
                {
                    firstRow = cell.Row;                 // Вычисляем верхний элемент
                    countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
                    endRow = firstRow + countRow;
                    do
                    {
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
                            if (dBConect.Value.CheckReadDB("SELECT * FROM base WHERE Article = '" + sArticle + "'"))
                            {
                                //Сборка запроса к БД
                                string commandText = String.Format($"INSERT INTO base (ip, klima, reserve, height, width, depth, article, execution, vendor)" +
                                      $" VALUES ('{sIP}', '{sKlima}', '{sReserve}', '{sHeinght}', '{sWidth}', '{sDepth}', '{sArticle}', '{sExecution}', 'None');");
                                //Оправка запроса к БД
                                dBConect.Value.MetodDB("SELECT * FROM base", commandText);
                                worksheet.get_Range("Z" + firstRow).Interior.ColorIndex = 0;

                                MessageBox.Show(
                                "Успешно записано в базу данных. Теперь доступна новая запись. \n Поздравляем!  \n" +
                                "Артикул = " + sArticle,
                                "Запись успешна!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                            }
                            else
                            {
                                MessageBox.Show(
                                "В базе данных уже есть такой артикул.\n Создавать новую запись не нужно. \n" +
                                "Артикул = " + sArticle,
                                "Ошибка записи!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                            }
                        }
                        else
                        {
                            MessageBox.Show(
                            "Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \n" +
                            "Артикул = " + sArticle,
                            "Ошибка записи",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.DefaultDesktopOnly);
                        }
                        firstRow++;
                    }
                    while (endRow > firstRow);
                }
                else
                {
                    MessageBox.Show(
                    "Программа работает только в файле " + dBConect.Value.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2) + "\n Пожайлуста откройте целевую книгу и запустите программу.",
                    "Ошибка вызова",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                }
                // Закрываем соединение с базой данных
                dBConect.Value.CloseDB();
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                exception.ToString(),
                "Ошибка надсройки",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
            }
        }
    }
}
