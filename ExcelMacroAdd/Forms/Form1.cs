using ExcelMacroAdd.Interfaces;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelMacroAdd.Forms
{
    internal partial class Form1 : Form
    {
        // Переменные иницализации                   
        private object confirmConversions = false;
        private object readOnly = false;
        private object addToRecentFiles = false;
        private object passwordDocument = Type.Missing;
        private object passwordTemplate = Type.Missing;
        private object revert = false;
        private object writePasswordDocument = Type.Missing;
        private object writePasswordTemplate = Type.Missing;
        private object format = Type.Missing;
        private object encoding = Type.Missing;
        private object oVisible = Type.Missing;
        private object openConflictDocument = Type.Missing;
        private object openAndRepair = Type.Missing;
        private object documentDirection = Type.Missing;
        private object noEncodingDialog = false;
        private object xmlTransform = Type.Missing;
        private object replaceTypeObj = Word.WdReplace.wdReplaceAll;
        private object wordMissing = Missing.Value;
        private readonly IResources resources;
        private readonly string PPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");

        internal Form1(IResources resources)
        {
            this.resources = resources;            
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
            Workbook workBook = Globals.ThisAddIn.GetActiveWorkBook();
            Excel.Range cell = Globals.ThisAddIn.GetActiveCell();

            int firstRow, countRow, endRow;

            firstRow = cell.Row;                 // Вычисляем верхний элемент
            countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
            endRow = firstRow + countRow;

            progressBar1.Minimum = 0;
            progressBar1.Maximum = countRow;
            progressBar1.Step = 1;

            new Thread(() =>
            {     
                int progressValue = 0;               
               //Инициализируем параметры Word
                Word.Application applicationWord = new Word.Application();
                // Переменная объект документа
                Word.Document document;               

                try
                {
                    // Цикл переборки строк
                    do
                    {
                        Object filename;
                        // Преобразование типов для определения формата сохранения 
                        if (int.TryParse(worksheet.Cells[firstRow, 14].Value2.ToString(), out int result) && result < resources.HeihgtMaxBox)
                        {
                            // переменная для открытия Word
                            filename = Path.Combine(PPath, resources.TempleteWall);
                        }
                        else
                        {
                            // переменная для открытия Word
                            filename = Path.Combine(PPath, resources.TempleteFloor);
                        }

                        string numberSave = Convert.ToString(worksheet.Cells[firstRow, 21].Value2);
                        string sTY = Convert.ToString(worksheet.Cells[firstRow, 8].Value2);
                        string sIcu = (Convert.ToString(worksheet.Cells[firstRow, 10].Value2));
                        string sIP = Convert.ToString(worksheet.Cells[firstRow, 11].Value2);
                        string sGab = (Convert.ToString(worksheet.Cells[firstRow, 14].Value2) + "x" + Convert.ToString(worksheet.Cells[firstRow, 15].Value2) +
                            "x" + Convert.ToString(worksheet.Cells[firstRow, 16].Value2));
                        string sMark = Convert.ToString(worksheet.Cells[firstRow, 4].Value2);
                        string sNum = Convert.ToString(worksheet.Cells[firstRow, 21].Value2);
                        string sKlima = Convert.ToString(worksheet.Cells[firstRow, 12].Value2);
                        string sUe = (Convert.ToString(worksheet.Cells[firstRow, 9].Value2));
                        string sGround = Convert.ToString(worksheet.Cells[firstRow, 28].Value2);
                        string sName = Convert.ToString(worksheet.Cells[firstRow, 6].Value2);
                        string sPaste = FuncReplece(sName ?? String.Empty); // ссылка на метод замены
                        string sZapol = Convert.ToString(worksheet.Cells[firstRow, 7].Value2);
                        string sSklon = FuncReplece(sZapol ?? String.Empty); // ссылка на метод замены
                        string sIsp = Convert.ToString(worksheet.Cells[firstRow, 27].Value2);
                        string sKorp = Convert.ToString(worksheet.Cells[firstRow, 29].Value2);
                        string nameFolderSafe = Convert.ToString(worksheet.Cells[firstRow, 1].Value2);

                        //Открываем Word
                        document = applicationWord.Documents.Open(filename, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate,
                        revert, writePasswordDocument, writePasswordTemplate, format, encoding, oVisible, openAndRepair, documentDirection,
                        noEncodingDialog, xmlTransform);
                        applicationWord.Visible = false;
                        //Инициализация метода Find
                        Word.Find find = applicationWord.Selection.Find;

                        // Замены ТУ
                        find.Execute("#ТУ", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sTY, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Ток
                        find.Execute("#Ток", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sIcu, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены IP
                        find.Execute("#IP", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sIP, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Габарит
                        find.Execute("#Габарит", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sGab, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Марка
                        find.Execute("#Марка", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sMark, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Номер
                        find.Execute("#Номер", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sNum, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Климат
                        find.Execute("#Климат", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sKlima, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Заземление
                        find.Execute("#Заземление", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sGround, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Название
                        find.Execute("#Название", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sName, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замена Вставка (необходим метод замены)
                        find.Execute("#Вставка", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sPaste, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Заполнение
                        find.Execute("#Заполнение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sZapol, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замена Склонение (необходим метод замены)
                        find.Execute("#Склонение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                        sSklon, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        // Замены Напряжение
                        if (sUe == "380")
                        {
                            find.Execute("#Напряжение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "~230/380 В.", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        else if (sIsp == "220")
                        {
                            find.Execute("#Напряжение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "~230В.", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        else
                        {
                            find.Execute("#Напряжение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            sUe + "В.", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        // Замены Исполнение
                        if (sIsp == "МП")
                        {
                            find.Execute("#Исполнение", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "монтажной плате", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        else if (sIsp == "ДР")
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
                        if (sKorp == "Металл")
                        {
                            find.Execute("#Корпус", ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing,
                            "металлическом", ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                        }
                        else if (sIsp == "ДР")
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
                        string folderName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Паспорта " + nameFolderSafe);
                        DirectoryInfo drInfo = new DirectoryInfo(folderName);
                        // Проверяем есть ли папка, если нет создаем
                        if (!drInfo.Exists)
                        {
                            drInfo.Create();
                            Logger(folderName);
                        }

                        document.SaveAs($@"{folderName}\Паспорт {numberSave}.docx");                        

                        int amountSheet = document.ComputeStatistics(WdStatistic.wdStatisticPages, false);
                        // Вызов логгера
                        Logger(folderName, numberSave, amountSheet);

                        document.Close();
                        firstRow++;

                        // Работа с элементами формы через делегат
                        this.Invoke((MethodInvoker)delegate ()
                        {
                            progressBar1.PerformStep();
                            label1.Text = $"Подождите пожайлуста, идет заполнение паспортов {++progressValue}/{countRow}.";
                        });
                    }
                    while (endRow > firstRow);

                    // Работа с элементами формы через делегат
                    this.Invoke((MethodInvoker)delegate ()
                    {
                        label1.Text = "Паспота заполнены. Ты молодец";
                        button1.Enabled = true;
                    });
                    applicationWord.Quit();
                }
                catch (COMException)
                {
                    MessageBox.Show(
                    "Проверьте имя проекта внимательно,\n экстрасенсы говорят что проблема в первом столбце.",
                    "Ошибка надстройки",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);

                    if (!(applicationWord is null)) applicationWord.Quit();
                }
                catch (RuntimeBinderException)
                {
                    MessageBox.Show(
                    "Проверьте заполнение всех столбцов,\n где-то нехватает данных.",
                    "Ошибка надстройки",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);

                    if (!(applicationWord is null)) applicationWord.Quit();
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

                    if (!(applicationWord is null)) applicationWord.Quit();
                }
            }).Start();
        }

        /// <summary>
        /// Метод для записи логов шапки документа
        /// </summary>
        /// <param name="folder"></param>
        private static void Logger(string folder)
        {
            string patch = Path.Combine(folder, "log.txt");
            using (StreamWriter output = File.AppendText(patch))
            {
                output.WriteLine("Версия OC:          " + Environment.OSVersion);
                output.WriteLine("Имя пользователя:   " + Environment.UserName);
                output.WriteLine("Имя компьютера:     " + Environment.MachineName);
                output.WriteLine("--------------------------------------------------------------------------------");
            }
        }

        /// <summary>
        /// Метод для записи логов формиррования паспортов
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="saveNum"></param>
        /// <param name="amount"></param>
        private static void Logger(string folder, string saveNum, int amount)
        {
            string patch = Path.Combine(folder, "log.txt");
            using (StreamWriter output = File.AppendText(patch))
            {
                output.WriteLine($"{DateTime.Now} | Паспорт {saveNum} сформирован успешно, в паспорте {amount} листа");
            }
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
