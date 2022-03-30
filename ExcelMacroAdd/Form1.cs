using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Threading;

namespace ExcelMacroAdd
{
    public partial class Form1 : Form
    {
        Object wordMissing = Missing.Value;
               
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {            
            var sync = SynchronizationContext.Current;

            Excel.Application application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet worksheet = ((Excel.Worksheet)application.ActiveSheet);
            Excel.Range cell = application.Selection;

            int firstRow, countRow, endRow;

            firstRow = cell.Row;                 // Вычисляем верхний элемент
            countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
            endRow = firstRow + countRow;

            progressBar1.Minimum = 0;
            progressBar1.Maximum = countRow;
            progressBar1.Step = 1;

            new Thread(() =>
            {
                var classDB = new DBConect();

                try
                {
                    // Открываем соединение с базой данных    
                    classDB.OpenDB();
               
                    int progressValue = 0;

                    int iHeihgtMax = Convert.ToInt32(classDB.RequestDB("SELECT * FROM settings WHERE set_name = 'sHeihgtMax';"));          // Запрашиваем максимальную высоту навесных шкафов
                                                                                                                                           //Инициализируем параметры Word
                    Word.Application applicationWord = new Word.Application();
                    // Переменная объект документа
                    Word.Document document;

                    // Переменные иницализации                   
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
                        // Ловим ошибку приведения типов
                        try
                        {
                            if ((Convert.ToInt32(worksheet.Cells[firstRow, 14].Value2) > iHeihgtMax))
                            {
                                filename = classDB.pPatch + classDB.RequestDB("SELECT * FROM settings WHERE set_name = 'sFloor';");
                            }
                            else
                            {
                                filename = classDB.pPatch + classDB.RequestDB("SELECT * FROM settings WHERE set_name = 'sWall';");
                            }
                        }
                        catch (FormatException)
                        {
                            filename = classDB.pPatch + classDB.RequestDB("SELECT * FROM settings WHERE set_name = 'sFloor';");
                        }

                        // переменная для имени сохраниея
                        string numberSave = Convert.ToString(worksheet.Cells[firstRow, 21].Value2);
                        string s_ty = Convert.ToString(worksheet.Cells[firstRow, 8].Value2);
                        string s_icu = (Convert.ToString(worksheet.Cells[firstRow, 10].Value2));
                        string s_ip = Convert.ToString(worksheet.Cells[firstRow, 11].Value2);
                        string s_gab = (Convert.ToString(worksheet.Cells[firstRow, 14].Value2) + "x" + Convert.ToString(worksheet.Cells[firstRow, 15].Value2) +
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
                            s_ue + "В.", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
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
                        string folderName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Паспорта " + folderSafe;
                        DirectoryInfo drInfo = new DirectoryInfo(folderName);
                        // Проверяем есть ли папка, если нет создаем
                        if (!drInfo.Exists)
                        {
                            drInfo.Create();
                        }
                        document.SaveAs(folderName + @"\Паспорт " + numberSave + ".docx");
                        document.Close();
                        firstRow++;

                        // Работа с элементами формы
                        sync.Post(__ => progressBar1.PerformStep(), null);
                        sync.Post(__ => label1.Text = "Подождите пожайлуста, идет заполнение паспортов " + ++progressValue + "/" + countRow, null);

                    }
                    while (endRow > firstRow);

                    // Работа с элементами формы
                    sync.Post(__ => label1.Text = "Паспота заполнены. Ты молодец", null);
                    sync.Post(__ => button1.Enabled = true, null);

                    classDB.CloseDB();
                    applicationWord.Quit();
                                       
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

            }).Start();
        }

        private string FuncReplece(string mReplase)                          // Функция замены
        {
            return mReplase.Replace("Щиток", "Щитка").Replace("Щит", "Щита").Replace("Шкаф", "Шкафа").Replace("Устройство", "Устройства").Replace("Корпус", "Корпуса").
                Replace("Ящик", "Ящика").Replace("Бокс", "Бокса").Replace("Панель", "Панели").Replace("распределительный", "распределительного");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close(); // Закрываем форму
        }
    }
}
