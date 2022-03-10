using Microsoft.Office.Tools.Ribbon;
using System;
using System.Data.OleDb;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelMacroAdd
{
    public partial class Ribbon1
    {
        // строка подключения к MS Access
        private OleDbConnection myConnection;
        //public string pPatch = @"\\192.168.100.100\ftp\Info_A\FTP\Производство Абиэлт\Инженеры\"; // Путь к базе данных
        public string pPatch = @"C:\Users\ПК\Desktop\Прайсы\Макро\";
        public string sPatch = "BdMacro.mdb";                                                     // Название файла базы данных
        Object wordMissing = Missing.Value;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {


        }
        
        private void button1_Click(object sender, RibbonControlEventArgs e) //Удаление формул
        {
            Excel.Application application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet worksheet = ((Excel.Worksheet)application.ActiveSheet);
            Excel.Range cell = application.Selection;
            cell.Value = cell.Value;                    //Удаляем формулы
            worksheet.get_Range("A1", Type.Missing).Select();   //Фокус на ячейку А1        
        }

        private void button2_Click(object sender, RibbonControlEventArgs e) //Разметка шаблона расчетов
        {
            Excel.Application application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet worksheet = ((Excel.Worksheet)application.ActiveSheet);        

            //состовляем надписи колонок           
            worksheet.get_Range("A1").Value2 = "Артикул";
            worksheet.get_Range("B1").Value2 = "Описание";
            worksheet.get_Range("C1").Value2 = "Кол-во";
            worksheet.get_Range("D1").Value2 = "Кратность";
            worksheet.get_Range("E1").Value2 = "Пр-ль";
            worksheet.get_Range("F1").Value2 = "Скидка";
            worksheet.get_Range("G1").Value2 = "Цена";
            worksheet.get_Range("H1").Value2 = "Цена со скидкой";
            worksheet.get_Range("I1").Value2 = "Стоимость";

            //увеличиваем размер по ширине диапазон ячеек
            worksheet.get_Range("A1", Type.Missing).EntireColumn.ColumnWidth = 21;
            worksheet.get_Range("B1", Type.Missing).EntireColumn.ColumnWidth = 80;
            worksheet.get_Range("C1", Type.Missing).EntireColumn.ColumnWidth = 10;
            worksheet.get_Range("D1", "I1").EntireColumn.ColumnWidth = 13;

            //размечаем границы и правим шрифты
            worksheet.get_Range("A1", "i500").Cells.Font.Name = "Calibri";
            worksheet.get_Range("A1", "i500").Cells.Font.Size = 11;

            var Excelcells = worksheet.get_Range("A1", "I11");
            Excel.XlBordersIndex BorderIndex;

            BorderIndex = Excel.XlBordersIndex.xlEdgeLeft; //Левая граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;

            BorderIndex = Excel.XlBordersIndex.xlEdgeTop; //Верхняя граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;

            BorderIndex = Excel.XlBordersIndex.xlEdgeBottom; //Нижняя граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;

            BorderIndex = Excel.XlBordersIndex.xlEdgeRight;  //Правая граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;

            BorderIndex = Excel.XlBordersIndex.xlInsideHorizontal;  //Внутренняя горизонтальня граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;

            BorderIndex = Excel.XlBordersIndex.xlInsideVertical;  //Внутренняя горизонтальня граница
            Excelcells.Borders[BorderIndex].Weight = Excel.XlBorderWeight.xlThin;
            Excelcells.Borders[BorderIndex].LineStyle = Excel.XlLineStyle.xlContinuous;
            Excelcells.Borders[BorderIndex].ColorIndex = 0;
        }

        private void button3_Click(object sender, RibbonControlEventArgs e) //Корпуса щитов
        {
            Excel.Application application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet worksheet = ((Excel.Worksheet)application.ActiveSheet);
            Excel.Range cell = application.Selection;
            int firstRow, countRow, endRow;

            try
            {
                if (application.ActiveWorkbook.Name == SettingsShow("sJornal"))            // Проверка по имени книги
                {
                    firstRow = cell.Row;                 // Вычисляем верхний элемент
                    countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
                    endRow = firstRow + countRow;

                    myConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + pPatch + sPatch + ";");
                    // открываем соединение с БД
                    myConnection.Open();
                    
                    do
                    {
                        string artNum = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);                                    
                        string query = "SELECT * FROM base WHERE article = '" + artNum + "'";
                        // Собираем запрос к БД
                        OleDbCommand command = new OleDbCommand(query, myConnection);
                        // Подсвечиваем ячейки с отсутствующими записями в БД
                        if (command.ExecuteScalar() == null)
                        {
                            worksheet.get_Range("Z" + firstRow).Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
                        }

                        OleDbDataReader reader = command.ExecuteReader();
                        // Заподние ячеек целевой книги
                        
                        while (reader.Read())
                        {                     
                            worksheet.get_Range("K" + firstRow).Value2 = reader[1].ToString();
                            worksheet.get_Range("L" + firstRow).Value2 = reader[3].ToString();
                            worksheet.get_Range("M" + firstRow).Value2 = reader[4].ToString();
                            worksheet.get_Range("N" + firstRow).Value2 = reader[2].ToString();
                            worksheet.get_Range("O" + firstRow).Value2 = reader[5].ToString();
                            worksheet.get_Range("P" + firstRow).Value2 = reader[6].ToString();
                            worksheet.get_Range("AC" + firstRow).Value2 = reader[8].ToString();
                        }
                        firstRow++;
                    }
                    while (endRow > firstRow);

                }
                else
                {
                    MessageBox.Show(
                    "Программа работает только в файле " + SettingsShow("sJornal") + "\n Пожайлуста откройте целевую книгу и запустите программу.",
                    "Ошибка вызова",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                }
            }

            catch (OleDbException exception)
            {
                MessageBox.Show(
                exception.ToString(),
                "Ошибка базы данных",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
            }

            finally
            {
                // Закрываем соединение с БД                
                myConnection.Dispose();
                myConnection.Close();
            }
        }

        private void button4_Click(object sender, RibbonControlEventArgs e) // Занесение в базу данных корпуса
        {
            Excel.Application application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet worksheet = ((Excel.Worksheet)application.ActiveSheet);
            Excel.Range cell = application.Selection;
            int firstRow, countRow, endRow;
            string sIP, sKlima, sReserve, sHeinght, sWidth, sDepth, sArticle, sExecution;

            try
            {
                if (application.ActiveWorkbook.Name == SettingsShow("sJornal"))            // Проверка по имени книги
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
                        if (!(sIP == null) && !(sKlima == null) && !(sReserve == null) && !(sHeinght == null)
                            && !(sWidth == null) && !(sDepth == null) && !(sArticle == null) && !(sExecution == null))                
                        {
                            myConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + pPatch + sPatch + ";");
                            // открываем соединение с БД
                            myConnection.Open();
                            string queryRead = "SELECT * FROM base WHERE Article = '" + sArticle + "'";
                            OleDbCommand commandRead = new OleDbCommand(queryRead, myConnection);
                            if (commandRead.ExecuteScalar() == null)
                            {                             
                                string queryInto = "SELECT * FROM base";
                                // Собираем запрос к БД
                                OleDbCommand commandInto = new OleDbCommand(queryInto, myConnection)
                                {
                                    Connection = myConnection,
                                    CommandText = "INSERT INTO base (ip, klima, reserve, height, width, depth, article, execution, vendor)" +
                                    " VALUES ('"+ sIP+"', '" + sKlima + "','"+ sReserve + "','"+ sHeinght +"','"+ sWidth +"','"+ sDepth +"','"+ sArticle +"','"+ sExecution +"','None')"
                                };
                                commandInto.ExecuteNonQuery();
                                // Освобождаем процессы
                                commandInto.Dispose(); 
                                commandRead.Dispose();
                                // Сбрасываем цвет ячейки на стандатрный
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
                                commandRead.Dispose();
                                
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
                    "Программа работает только в файле _Журнал учета НКУ 2022.xlsx \n Пожайлуста откройте целевую книгу и запустите программу.",
                    "Ошибка вызова",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                }
            }
            catch (OleDbException exception)
            {
                MessageBox.Show(
                exception.ToString(),
                "Ошибка базы данных",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
                throw;
            }

            finally
            {
                // Закрываем соединение с БД
                myConnection.Dispose();
                myConnection.Close();
            }
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)  // Заполнение паспортов
        {
            Excel.Application application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet worksheet = ((Excel.Worksheet)application.ActiveSheet);
            Excel.Range cell = application.Selection;
            int firstRow, countRow, endRow;

            firstRow = cell.Row;                 // Вычисляем верхний элемент
            countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
            endRow = firstRow + countRow;

            try
            {
                if (application.ActiveWorkbook.Name == SettingsShow("sJornal"))            // Проверка по имени книги
                {
                    int iHeihgtMax = Convert.ToInt32(SettingsShow("sHeihgtMax"));          // Запрашиваем максимальную высоту навесных шкафов
                    //Инициализируем параметры Word
                    Word.Application applicationWord = new Word.Application();
                    // Переменная объект документа
                    Word.Document document;

                    // Переменные иницализации
                    //Object filename = pPatch + "Паспорт_напольные.docx";
                    Object confirmConversions = false;
                    Object readOnly = false;
                    Object addToRecentFiles = false;
                    Object passwordDocument = Type.Missing;
                    Object passwordTemplate = Type.Missing;
                    Object revert = false;
                    Object writePasswordDocument = Type.Missing;
                    Object writePasswordTemplate = Type.Missing;
                    Object format = Type.Missing;
                    Object encoding = Type.Missing;
                    Object oVisible = Type.Missing;
                    Object openConflictDocument = Type.Missing;
                    Object openAndRepair = Type.Missing;
                    Object documentDirection = Type.Missing;
                    Object noEncodingDialog = false;
                    Object xmlTransform = Type.Missing;
                    Object replaceTypeObj = Word.WdReplace.wdReplaceAll;

                    // Цикл переборки строк
                    do
                    {
                        Object filename;
                        if ((Convert.ToInt32(worksheet.Cells[firstRow, 14].Value2) > iHeihgtMax))
                        {
                            filename = pPatch + SettingsShow("sFloor");
                        }
                        else
                        {
                            filename = pPatch + SettingsShow("sWall");
                        }

                        // переменная для имени сохраниея
                        string numberSave = Convert.ToString(worksheet.Cells[firstRow, 21].Value2);
                        string s_ty = Convert.ToString(worksheet.Cells[firstRow, 8].Value2);
                        string s_icu = (Convert.ToString(worksheet.Cells[firstRow, 10].Value2));
                        string s_ip = Convert.ToString(worksheet.Cells[firstRow, 11].Value2);
                        string s_gab = (Convert.ToString(worksheet.Cells[firstRow, 14].Value2) + "x" +  Convert.ToString(worksheet.Cells[firstRow, 15].Value2) +
                            "x" + Convert.ToString(worksheet.Cells[firstRow, 16].Value2));
                        string s_mark = Convert.ToString(worksheet.Cells[firstRow, 4].Value2);
                        string s_num = Convert.ToString(worksheet.Cells[firstRow, 21].Value2);
                        string s_klima = Convert.ToString(worksheet.Cells[firstRow, 12].Value2);
                        string s_ue = (Convert.ToString(worksheet.Cells[firstRow, 9].Value2));
                        string s_ground = Convert.ToString(worksheet.Cells[firstRow, 28].Value2);
                        string s_name = Convert.ToString(worksheet.Cells[firstRow, 6].Value2);
                        string s_paste = FuncReplece(s_name); // ссылка на метод замены
                        string s_zapol = Convert.ToString(worksheet.Cells[firstRow, 7].Value2);
                        string s_slon = FuncReplece(s_zapol); // ссылка на метод замены
                        string s_isp = Convert.ToString(worksheet.Cells[firstRow, 27].Value2);
                        string s_korp = Convert.ToString(worksheet.Cells[firstRow, 29].Value2);
                        string folderSafe = Convert.ToString(worksheet.Cells[firstRow, 1].Value2);                       

                        //Открываем Word
                        document = applicationWord.Documents.Open(filename, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate,
                        revert, writePasswordDocument, writePasswordTemplate, format, encoding, oVisible, openAndRepair, documentDirection,
                        noEncodingDialog, xmlTransform);
                        applicationWord.Visible = false;
                        //Инициализация метода Find
                        Word.Find find = applicationWord.Selection.Find;

                        // Замены ТУ
                        find.Execute("#ТУ", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_ty, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Ток
                        find.Execute("#Ток", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_icu, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены IP
                        find.Execute("#IP", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_ip, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Габарит
                        find.Execute("#Габарит", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_gab, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Марка
                        find.Execute("#Марка", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_mark, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Номер
                        find.Execute("#Номер", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_num, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Климат
                        find.Execute("#Климат", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_klima, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Заземление
                        find.Execute("#Заземление", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_ground, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Название
                        find.Execute("#Название", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_name, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замена Вставка (необходим метод замены)
                        find.Execute("#Вставка", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_paste, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Заполнение
                        find.Execute("#Заполнение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_zapol, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замена Склонение (необходим метод замены)
                        find.Execute("#Склонение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        s_slon, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Напряжение
                        if (s_ue == "380")
                        {
                            find.Execute("#Напряжение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "~230/380 В.", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        else if (s_isp == "220")
                        {
                            find.Execute("#Напряжение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "~230В.", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        else
                        {
                            find.Execute("#Напряжение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            s_ue +"В.", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        // Замены Исполнение
                        if (s_isp == "МП")
                        {
                            find.Execute("#Исполнение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "монтажной плате", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        else if (s_isp == "ДР")
                        {
                            find.Execute("#Исполнение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "din-рейках", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        else
                        {
                            find.Execute("#Исполнение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "монтажной плате, din-рейках", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        // Замены Материал
                        if (s_korp == "Металл")
                        {
                            find.Execute("#Корпус", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "металлическом", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        else if (s_isp == "ДР")
                        {
                            find.Execute("#Корпус", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "пластиковом", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        else
                        {
                            find.Execute("#Корпус", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "металлическом или пластиковом", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }

                        //Путь к папке Рабочего стола                                     
                        string folderName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +  @"\Паспорта " + folderSafe;
                        DirectoryInfo drInfo = new DirectoryInfo(folderName);
                        // Проверяем есть ли папка, если нет создаем
                        if (!drInfo.Exists)
                        {
                            drInfo.Create();
                        }  
                        document.SaveAs(folderName +  @"\Паспорт" + numberSave + ".docx");
                        document.Close();          
                        firstRow++;
                    }
                    while (endRow > firstRow);
                    applicationWord.Quit();
                }
                else
                {
                    MessageBox.Show(
                    "Программа работает только в файле " + SettingsShow("sJornal") + "\n Пожайлуста откройте целевую книгу и запустите программу.",
                    "Ошибка вызова",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
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

        private void button7_Click(object sender, RibbonControlEventArgs e)  // "Прическа" расчетов
        {
            Excel.Application application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook workBook = (Excel.Workbook)application.ActiveWorkbook;
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

        private string SettingsShow (string sSettings)                       //База данных настроек
        {
            string retSett = null;
            try
            {
                myConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + pPatch + sPatch + ";");
                // открываем соединение с БД
                myConnection.Open();

                string query = "SELECT * FROM settings WHERE set_name = '" + sSettings + "'";
                // Собираем запрос к БД
                OleDbCommand command = new OleDbCommand(query, myConnection);

                OleDbDataReader reader = command.ExecuteReader();
                // Заподние ячеек целевой книги
                while (reader.Read())
                {
                    retSett = (reader[2].ToString());
                }   
                return retSett;
            }

            catch (OleDbException exception)
            {
                MessageBox.Show(
                exception.ToString(),
                "Ошибка базы данных",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
                return null;
            }
            finally
            {
                myConnection.Dispose();
                myConnection.Close();
            }
        }

        private string FuncReplece(string mReplase)
        {
            return mReplase.Replace("Щиток", "Щитка").Replace("Щит", "Щита").Replace("Шкаф", "Шкафа").Replace("Устройство", "Устройства").Replace("Корпус", "Корпуса").
                Replace("Ящик", "Ящика").Replace("Бокс", "Бокса").Replace("Панель", "Панели").Replace("распределительный", "распределительного");
        }

        private async void button8_Click(object sender, RibbonControlEventArgs e)
        {
            await Task.Run(() =>
            {
                Form1 fs = new Form1();
                fs.ShowDialog();
                Thread.Sleep(5000);
            });
        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet worksheet = ((Excel.Worksheet)application.ActiveSheet);
            Excel.Range cell = application.Selection;
            int firstRow;
            string sIP, sKlima, sReserve, sHeinght, sWidth, sDepth, sArticle, sExecution;

            try
            {
                if (application.ActiveWorkbook.Name == SettingsShow("sJornal"))            // Проверка по имени книги
                {

                    firstRow = cell.Row;                 // Вычисляем верхний элемент
                    sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);
                    myConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + pPatch + sPatch + ";");
                    // открываем соединение с БД
                    myConnection.Open();
                    string queryRead = "SELECT * FROM base WHERE Article = '" + sArticle + "'";
                    OleDbCommand commandRead = new OleDbCommand(queryRead, myConnection);
                    if (!(commandRead.ExecuteScalar() == null))
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
                        if (!(sIP == null) && !(sKlima == null) && !(sReserve == null) && !(sHeinght == null)
                            && !(sWidth == null) && !(sDepth == null) && !(sArticle == null) && !(sExecution == null))
                        {
                            string queryUpdate = "SELECT * FROM base";
                            // Собираем запрос к БД
            
                            OleDbCommand commandUpdate = new OleDbCommand(queryUpdate, myConnection)
                            {
                                Connection = myConnection,
                                // CommandText = "UPDATE base SET (ip, klima, reserve, height, width, depth, execution, vendor)" +
                                // " VALUES ('" + sIP + "', '" + sKlima + "','" + sReserve + "','" + sHeinght + "','" + sWidth + "','" + sDepth + "','" + sExecution + "','None') WHERE Article = '" + sArticle + "'"


                                CommandText = "UPDATE base SET ip = '"+ sIP + "' WHERE Article = '" +sArticle +"'"

                            };
                            commandUpdate.ExecuteNonQuery();
                            // Освобождаем процессы
                            commandUpdate.Dispose();
                            commandRead.Dispose();
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
                    }
                    else
                    {
                        commandRead.Dispose();

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
            catch (OleDbException exception)
            {
                MessageBox.Show(
                exception.ToString(),
                "Ошибка базы данных",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
            }
            finally
            {
                // Закрываем соединение с БД
                myConnection.Dispose();
                myConnection.Close();
            }
                    }

    }
}
