using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd
{
    public partial class Ribbon1
    {     
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            GetValuteTSB getRate = new GetValuteTSB();
            this.label1.Label = "Доллар = " + getRate.USDRate;
            this.label2.Label = "ЕВРО     = " + getRate.EvroRate;
            this.label3.Label = "Юань    = " + getRate.CnyRate;
        }
        private void button1_Click(object sender, RibbonControlEventArgs e) //Удаление формул
        {
            Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
            Range cell = Globals.ThisAddIn.GetActiveCell();

            cell.Value = cell.Value;                            //Удаляем формулы
            worksheet.get_Range("A1", Type.Missing).Select();   //Фокус на ячейку А1        
        }

        private void button2_Click(object sender, RibbonControlEventArgs e) => _ = new Linker(); //Разметка шаблона расчетов
 
        private void button3_Click(object sender, RibbonControlEventArgs e) //Корпуса щитов
        {        
            Excel.Application application = Globals.ThisAddIn.GetApplication();
            Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
            Range cell = Globals.ThisAddIn.GetActiveCell();

            int firstRow, countRow, endRow;
            // Создаем экземпляр класса DBConect
            var classDB = new DBConect();
            try
            {
                // Открываем соединение с базой данных    
                classDB.OpenDB();

                if (application.ActiveWorkbook.Name == classDB.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';" ,2))            // Проверка по имени книги
                {
                    firstRow = cell.Row;                 // Вычисляем верхний элемент
                    countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
                    endRow = firstRow + countRow;
                    // Инициализируем структуру для записи                
                    DBtable dBtable = new DBtable();  
                    do
                    {
                        string sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);
                        string query = "SELECT * FROM base WHERE article = '" + sArticle + "'";

                        if (classDB.CheckReadDB(query))
                        {
                            worksheet.get_Range("Z" + firstRow).Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
                        }
                        else
                        {
                            // Передеем структуру по референсной ссылке в библиотечный метод 
                            classDB.ReadingDB(query,ref dBtable);                  
                            // Присваеваем ячейкам данные из массива
                            worksheet.get_Range("K" + firstRow).Value2  = dBtable.ipTable        ?? String.Empty;
                            worksheet.get_Range("L" + firstRow).Value2  = dBtable.klimaTable     ?? String.Empty;
                            worksheet.get_Range("M" + firstRow).Value2  = dBtable.reserveTable   ?? String.Empty;
                            worksheet.get_Range("N" + firstRow).Value2  = dBtable.heightTable    ?? String.Empty;
                            worksheet.get_Range("O" + firstRow).Value2  = dBtable.widthTable     ?? String.Empty;
                            worksheet.get_Range("P" + firstRow).Value2  = dBtable.depthTable     ?? String.Empty;
                            worksheet.get_Range("AC" + firstRow).Value2 = dBtable.executionTable ?? String.Empty;
                        }                      
                        
                        firstRow++;
                    }
                    while (endRow > firstRow);

                }
                else
                {
                    MessageBox.Show(                    
                    "Программа работает только в файле " + classDB.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2) + "\n Пожайлуста откройте целевую книгу и запустите программу.",
                    "Ошибка вызова",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                }

                // Закрываем соединение с базой данных
                classDB.CloseDB();
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
        private void button4_Click(object sender, RibbonControlEventArgs e) // Занесение в базу данных корпуса
        {
            Excel.Application application = Globals.ThisAddIn.GetApplication();
            Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
            Range cell = Globals.ThisAddIn.GetActiveCell();

            int firstRow, countRow, endRow;
            string sIP, sKlima, sReserve, sHeinght, sWidth, sDepth, sArticle, sExecution;
                      
            var classDB = new DBConect();                // Создаем экземпляр класса
            try
            {
                // Открываем соединение с базой данных    
                classDB.OpenDB();
                // Проверка по имени книги
                if (application.ActiveWorkbook.Name == classDB.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2))           
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
                            if (classDB.CheckReadDB("SELECT * FROM base WHERE Article = '" + sArticle + "'"))
                                { 
                                string commandText = "INSERT INTO base (ip, klima, reserve, height, width, depth, article, execution, vendor)" +
                                      " VALUES ('" + sIP + "', '" + sKlima + "','" + sReserve + "','" + sHeinght + "','" + sWidth + "','" + sDepth + "','" + sArticle + "','" + sExecution + "','None');";

                                classDB.MetodDB("SELECT * FROM base", commandText);
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
                            "Артикул = "+ sArticle,
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
                    "Программа работает только в файле " + classDB.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2) + "\n Пожайлуста откройте целевую книгу и запустите программу.",
                    "Ошибка вызова",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                }

                // Закрываем соединение с базой данных
                classDB.CloseDB();

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

        private void button6_Click(object sender, RibbonControlEventArgs e)  // Заполнение паспортов
        {
            Excel.Application application = Globals.ThisAddIn.GetApplication();

            var classDB = new DBConect();
            classDB.OpenDB();
            // Проверка по имени книги
            if (application.ActiveWorkbook.Name == classDB.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2))
            {
                OpenForm();
            }
            else
            {
                MessageBox.Show(
                "Программа работает только в файле " + classDB.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2) + "\n Пожайлуста откройте целевую книгу и запустите программу.",
                "Ошибка вызова",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
            }
            classDB.CloseDB();
        }


        private async void OpenForm()
        {
            await Task.Run(() =>
            {
                Form1 fs = new Form1();
                fs.ShowDialog();
                Thread.Sleep(100);
            });
        }


        private void button7_Click(object sender, RibbonControlEventArgs e)  // "Прическа" расчетов
        {
            Workbook workBook = Globals.ThisAddIn.GetActiveWorkBook();

            foreach (Excel.Worksheet sheet in workBook.Sheets)
            {
                sheet.Activate();
                if (!(sheet.Index == 1))
                {                  
                    sheet.get_Range("A1", "i500").Cells.Font.Name = "Calibri";
                    sheet.get_Range("A1", "i500").Cells.Font.Size = 11;
                    sheet.get_Range("D1", Type.Missing).EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                    sheet.get_Range("D1", Type.Missing).Value2 = "Кратность";
                    sheet.get_Range("D1", Type.Missing).EntireColumn.ColumnWidth = 10;
                }             
            }
        }                  
              
        private void button9_Click(object sender, RibbonControlEventArgs e)   // Корректировка записей БД
        {

            DialogResult dialogResult = MessageBox.Show("Вы уверены, что хотите изменить запись в БД? \nИзменения коснуться всех пользователей.", 
                                                        "Контрольный вопрос", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {        
                Excel.Application application = Globals.ThisAddIn.GetApplication();
                Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
                Range cell = Globals.ThisAddIn.GetActiveCell();

                int firstRow;
                string sIP, sKlima, sReserve, sHeinght, sWidth, sDepth, sArticle, sExecution;
                var classDB = new DBConect();
                try
                {
                    // Открываем соединение с базой данных    
                    classDB.OpenDB();

                    if (application.ActiveWorkbook.Name == classDB.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2))            // Проверка по имени книги
                    {
                        firstRow = cell.Row;                 // Вычисляем верхний элемент
                        sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);

                        if (!(classDB.CheckReadDB("SELECT * FROM base WHERE Article = '" + sArticle + "'")))
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
                                string queryUpdate = "SELECT * FROM base";
                                // Собираем запрос к БД   
                                string data = @"UPDATE base " +
                                                    "  SET ip         = '" + sIP + "'," +
                                                    "      klima      = '" + sKlima + "'," +
                                                    "      reserve    = '" + sReserve + "'," +
                                                    "      height     = '" + sHeinght + "'," +
                                                    "      width      = '" + sWidth + "'," +
                                                    "      depth      = '" + sDepth + "'," +
                                                    "      execution  = '" + sExecution + "'" +
                                                    "WHERE article    = '" + sArticle + "'";

                                classDB.MetodDB(queryUpdate, data);   
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
                            classDB.CloseDB();

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

        /// <summary>
        /// Удаление формул на всех листах кроме первого
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook workBook = Globals.ThisAddIn.GetActiveWorkBook();

            foreach (Excel.Worksheet sheet in workBook.Sheets)
            {
               sheet.Activate();
               if (!(sheet.Index == 1))
               {                   
                    sheet.get_Range("A2", "G500").Value = sheet.get_Range("A2", "G500").Value;
                    sheet.get_Range("A1", Type.Missing).Select();   //Фокус на ячейку А1
                }
            }
        }

        /// <summary>
        /// Запуск "О программе" в отдельном процессе
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void button10_Click(object sender, RibbonControlEventArgs e)  // Открывает "О программе"
        {
            await Task.Run(() =>
            {
                AboutBox1 about = new AboutBox1();
                about.ShowDialog();
                Thread.Sleep(5000);            
            });
        }

        private async void button11_Click(object sender, RibbonControlEventArgs e)
        {
            await Task.Run(() =>
            {
                Form2 fs = new Form2();
                fs.ShowDialog();
                Thread.Sleep(5000);
            });    
        }

        private async void button12_Click(object sender, RibbonControlEventArgs e)
        {
            await Task.Run(() =>
            {
                Form3 fs = new Form3();
                fs.ShowDialog();
                Thread.Sleep(5000);
            });
        }
    }

}
