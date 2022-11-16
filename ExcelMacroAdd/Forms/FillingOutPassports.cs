using ExcelMacroAdd.Forms.FillingOutPassportClass;
using ExcelMacroAdd.Interfaces;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelMacroAdd.Forms
{
    internal partial class FillingOutPassports : Form
    {
        // Переменные иницализации                   
        private readonly object confirmConversions = false;
        private readonly object readOnly = false;
        private readonly object addToRecentFiles = false;
        private readonly object passwordDocument = Type.Missing;
        private readonly object passwordTemplate = Type.Missing;
        private readonly object revert = false;
        private readonly object writePasswordDocument = Type.Missing;
        private readonly object writePasswordTemplate = Type.Missing;
        private readonly object format = Type.Missing;
        private readonly object encoding = Type.Missing;
        private readonly object oVisible = Type.Missing;
        private readonly object openAndRepair = Type.Missing;
        private readonly object documentDirection = Type.Missing;
        private readonly object noEncodingDialog = false;
        private readonly object xmlTransform = Type.Missing;
        private static object _replaceTypeObj = WdReplace.wdReplaceAll;
        private static object _wordMissing = Missing.Value;
        private readonly IResources resources;
        private readonly string pPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template");

        readonly Func<Find, string, string, bool> replacingTextLabels = (find, label, value) => find.Execute(label, ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing,
            value, ref _replaceTypeObj, ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing);

        internal FillingOutPassports(IResources resources)
        {
            this.resources = resources;            
            InitializeComponent();
        }

        private void FillingOutPassports_Load(object sender, EventArgs e)
        {
            Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
            Excel.Range cell = Globals.ThisAddIn.GetActiveCell();

            var firstRow = cell.Row; // Вычисляем верхний элемент
            var countRow = cell.Rows.Count; // Вычисляем кол-во выделенных строк
            var endRow = firstRow + countRow;

            progressBar1.Minimum = 0;
            progressBar1.Maximum = countRow;
            progressBar1.Step = 1;

            new Thread(() =>
            {     
                int progressValue = 0;               
               //Инициализируем параметры Word
                Word.Application applicationWord = new Word.Application();
                // Переменная объект документа

                try
                {
                    // Цикл переборки строк
                    do
                    {
                        object filename;
                        // Преобразование типов для определения формата сохранения 
                        if (int.TryParse(worksheet.Cells[firstRow, 14].Value2.ToString(), out int result) && result < resources.HeightMaxBox)
                        {
                            // переменная для открытия Word
                            filename = Path.Combine(pPath, resources.TemplateWall);
                        }
                        else
                        {
                            // переменная для открытия Word
                            filename = Path.Combine(pPath, resources.TemplateFloor);
                        }

                        string numberSave = worksheet.Cells[firstRow, 21].Value2.ToString();
                        string sTy = worksheet.Cells[firstRow, 8].Value2.ToString();
                        string sIcu = worksheet.Cells[firstRow, 10].Value2.ToString();
                        string sIp = worksheet.Cells[firstRow, 11].Value2.ToString();
                        string sGab = worksheet.Cells[firstRow, 14].Value2.ToString() + "x"
                                    + worksheet.Cells[firstRow, 15].Value2.ToString() + "x"
                                    + worksheet.Cells[firstRow, 16].Value2.ToString();
                        string sMark = worksheet.Cells[firstRow, 4].Value2.ToString();
                        string sNum = worksheet.Cells[firstRow, 21].Value2.ToString();
                        //string sClimate = worksheet.Cells[firstRow, 12].Value2.ToString();
                        string sUe = worksheet.Cells[firstRow, 9].Value2.ToString();
                        string sGround = worksheet.Cells[firstRow, 28].Value2.ToString();
                        string sName = worksheet.Cells[firstRow, 6].Value2.ToString();
                        string sPaste = FuncReplace(sName ?? string.Empty); // ссылка на метод замены
                        string firstWords = worksheet.Cells[firstRow, 7].Value2.ToString();
                        string secondWords = FuncReplace(firstWords ?? string.Empty); // ссылка на метод замены
                        string sIsp = worksheet.Cells[firstRow, 27].Value2.ToString();
                        string sMaterial = worksheet.Cells[firstRow, 29].Value2.ToString();
                        string nameFolderSafe = worksheet.Cells[firstRow, 1].Value2.ToString();

                        //Открываем Word
                        var document = applicationWord.Documents.Open(filename, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate,
                            revert, writePasswordDocument, writePasswordTemplate, format, encoding, oVisible, openAndRepair, documentDirection,
                            noEncodingDialog, xmlTransform);
                        applicationWord.Visible = false;
                        //Инициализация метода Find
                        Find find = applicationWord.Selection.Find;

                        var verifyIsNotFind = new List<bool>(15);

                        // Замены ТУ
                        verifyIsNotFind.Add(
                            replacingTextLabels(find, "#ТУ", sTy));
                        // Замены Ток
                        verifyIsNotFind.Add(
                            replacingTextLabels(find, "#Ток", sIcu));
                        // Замены IP
                        verifyIsNotFind.Add(
                            replacingTextLabels(find, "#IP", sIp));
                        // Замены Габарит
                        verifyIsNotFind.Add(
                            replacingTextLabels(find, "#Габарит", sGab));
                        // Замены Марка
                        verifyIsNotFind.Add(
                            replacingTextLabels(find, "#Марка", sMark));
                        // Замены Номер
                        verifyIsNotFind.Add(
                            replacingTextLabels(find, "#Номер", sNum));
                        // Замены Климат
                        //verifyIsNotFind.Add(
                        //replacingTextLabels(find, "#Климат", sClimate));
                        // Замены Заземление
                        verifyIsNotFind.Add(
                            replacingTextLabels(find, "#Заземление", sGround));
                        // Замены Название
                        verifyIsNotFind.Add(
                            replacingTextLabels(find, "#Название", sName));
                        // Замена Вставка (необходим метод замены)
                        verifyIsNotFind.Add(
                            replacingTextLabels(find, "#Вставка", sPaste));
                        // Замены Заполнение
                        verifyIsNotFind.Add(
                            replacingTextLabels(find, "#Заполнение", firstWords));
                        // Замена Склонение (необходим метод замены)
                        verifyIsNotFind.Add(
                            replacingTextLabels(find, "#Склонение", secondWords));

                        // Замены Напряжение
                        if (sUe == "380")
                        {
                            verifyIsNotFind.Add(
                                replacingTextLabels(find, "#Напряжение", "~230/380 В."));
                        }
                        else if (sIsp == "220")
                        {
                            verifyIsNotFind.Add(
                                replacingTextLabels(find, "#Напряжение", "~230В."));
                        }
                        else
                        {
                            verifyIsNotFind.Add(
                                replacingTextLabels(find, "#Напряжение", sUe + "В."));
                        }

                        // Замены Исполнение
                        if (sIsp == "МП")
                        {
                            verifyIsNotFind.Add(
                                replacingTextLabels(find, "#Исполнение", "монтажной плате"));
                        }
                        else if (sIsp == "ДР")
                        {
                            verifyIsNotFind.Add(
                                replacingTextLabels(find, "#Исполнение", "din-рейках"));
                        }
                        else
                        {
                            verifyIsNotFind.Add(
                                replacingTextLabels(find, "#Исполнение", "монтажной плате, din-рейках"));
                        }

                        // Замены Материал
                        if (sMaterial == "Металл")
                        {
                            verifyIsNotFind.Add(
                                replacingTextLabels(find, "#Корпус", "металлическом"));
                        }
                        else if (sIsp == "ДР")
                        {
                            verifyIsNotFind.Add(
                                replacingTextLabels(find, "#Корпус", "пластиковом"));
                        }
                        else
                        {
                            verifyIsNotFind.Add(
                                replacingTextLabels(find, "#Корпус", "металлическом или пластиковом"));
                        }

                        //Путь к папке Рабочего стола                                     
                        string folderName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Паспорта " + nameFolderSafe);
                        DirectoryInfo drInfo = new DirectoryInfo(folderName);
                        // Проверяем есть ли папка, если нет создаем
                        if (!drInfo.Exists)
                        {
                            drInfo.Create();
                            WriteLog.Logger(folderName);
                        }

                        document.SaveAs($@"{folderName}\Паспорт {numberSave}.docx");                        

                        int amountSheet = document.ComputeStatistics(WdStatistic.wdStatisticPages, false);
                        
                        // Вызов логгера
                        WriteLog.Logger(folderName, numberSave, amountSheet);

                        document.Close();
                        firstRow++;

                        if (!verifyIsNotFind.All(item => item))
                            МеssageView.MessageWarning(
                                $@"В шаблоне не произошла вставка{Environment.NewLine} одного или нескольких значений.",
                                @"Проблема шаблона");

                        // Работа с элементами формы через делегат
                        this.Invoke((MethodInvoker)delegate
                        {
                            progressBar1.PerformStep();
                            label1.Text = $@"Подождите пожайлуста, идет заполнение паспортов {++progressValue}/{countRow}.";
                        });
                    }
                    while (endRow > firstRow);

                }
                catch (COMException)
                {
                    МеssageView.MessageError(
                        $@"Проверьте имя проекта внимательно,{Environment.NewLine} экстрасенсы говорят что проблема в первом столбце.",
                        @"Ошибка надстройки");
                    if (!(applicationWord is null)) applicationWord.Quit();
                }
                catch (RuntimeBinderException)
                {
                    МеssageView.MessageError(                                            
                        $@"Проверьте заполнение всех столбцов,{Environment.NewLine} где-то нехватает данных.",
                        @"Ошибка надстройки");
                    if (!(applicationWord is null)) applicationWord.Quit();
                }
                catch (Exception exception)
                {
                    МеssageView.MessageError(
                        exception.ToString(),
                        @"Ошибка надстройки");
                    if (!(applicationWord is null)) applicationWord.Quit();
                }
                finally
                {
                    // Работа с элементами формы через делегат
                    this.Invoke((MethodInvoker)delegate
                    {
                        label1.Text = @"Паспота заполнены. Ты молодец";
                        button1.Enabled = true;
                    });
                    applicationWord.Quit();
                }
            }).Start();
        }

        private string FuncReplace(string mReplase)                          // Функция замены
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
