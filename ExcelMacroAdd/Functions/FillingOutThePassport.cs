using ExcelMacroAdd.Forms;
using ExcelMacroAdd.Forms.Services;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelMacroAdd.Functions
{
    internal sealed class FillingOutThePassport : AbstractFunctions
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
        private readonly IFillingOutThePassportSettings resources;
        private readonly string pPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template");

        readonly Func<Find, string, string, bool> replacingTextLabels = (find, label, value) => find.Execute(label, ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing,
            ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing,  value, ref _replaceTypeObj, ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing);

        public delegate void MetodProgressStep(int step);
        public event MetodProgressStep ProgressStep;

        public delegate void MetodProgressFinal();
        public event MetodProgressFinal ProgressFinal;

        public FillingOutThePassport(IFillingOutThePassportSettings resources)
        {
            this.resources = resources;
        }

        public override void Start()
        {
            if (Application.ActiveWorkbook.Name != resources.NameFileJournal) // Проверка по имени книги
            {
                MessageWarning("Функция работает только в \"Журнале учета НКУ\" текущего года. \n Пожайлуста откройте необходимую книгу Excel.",
                            "Имя книги не совпадает с целевой");
                return;
            }

            var firstRow = Cell.Row; // Вычисляем верхний элемент
            var countRow = Cell.Rows.Count; // Вычисляем кол-во выделенных строк
            var endRow = firstRow + countRow;
            int numberStep = 0;

            new Thread(() =>
            {
                FillingOutPassports fs = new FillingOutPassports(countRow);
                ProgressStep += fs.OnStep;
                ProgressFinal += fs.OnFinal;
                fs.ShowDialog();
                Thread.Sleep(500);
            }).Start();

            new Thread(() =>
            {
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
                        if (int.TryParse(Worksheet.Cells[firstRow, 14].Value2.ToString(), out int result) && result < resources.HeightMaxBox)
                        {
                            // переменная для открытия Word
                            filename = Path.Combine(pPath, resources.TemplateWall);
                        }
                        else
                        {
                            // переменная для открытия Word
                            filename = Path.Combine(pPath, resources.TemplateFloor);
                        }

                        string factoryNumber = Worksheet.Cells[firstRow, 21].Value2.ToString();
                        string sTy = Worksheet.Cells[firstRow, 8].Value2.ToString();
                        string sIcu = Worksheet.Cells[firstRow, 10].Value2.ToString();
                        string sIp = Worksheet.Cells[firstRow, 11].Value2.ToString();
                        string sGab = Worksheet.Cells[firstRow, 14].Value2.ToString() + "x"
                                    + Worksheet.Cells[firstRow, 15].Value2.ToString() + "x"
                                    + Worksheet.Cells[firstRow, 16].Value2.ToString();
                        string sMark = Worksheet.Cells[firstRow, 4].Value2.ToString();
                        string sNum = Worksheet.Cells[firstRow, 21].Value2.ToString();
                        //string sClimate = worksheet.Cells[firstRow, 12].Value2.ToString();
                        string sUe = Worksheet.Cells[firstRow, 9].Value2.ToString();
                        string sGround = Worksheet.Cells[firstRow, 29].Value2.ToString();
                        string sName = Worksheet.Cells[firstRow, 6].Value2.ToString();
                        string sPaste = FuncReplace(sName ?? string.Empty); // ссылка на метод замены
                        string firstWords = Worksheet.Cells[firstRow, 7].Value2.ToString();
                        string secondWords = FuncReplace(firstWords ?? string.Empty); // ссылка на метод замены
                        string sIsp = Worksheet.Cells[firstRow, 27].Value2.ToString();
                        string sMaterial = Worksheet.Cells[firstRow, 30].Value2.ToString();
                        string sExecution = Worksheet.Cells[firstRow, 28].Value2.ToString();
                        string nameFolderSafe = Worksheet.Cells[firstRow, 1].Value2.ToString();

                        //Открываем Word
                        var document = applicationWord.Documents.Open(filename, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate,
                            revert, writePasswordDocument, writePasswordTemplate, format, encoding, oVisible, openAndRepair, documentDirection,
                            noEncodingDialog, xmlTransform);
                        applicationWord.Visible = false;
                        //Инициализация метода Find
                        Find find = applicationWord.Selection.Find;

                        var verifyIsNotFind = new List<bool>(15)
                        {
                            // Замены ТУ
                            replacingTextLabels(find, "#ТУ", sTy),
                            // Замены Ток
                            replacingTextLabels(find, "#Ток", sIcu),
                            // Замены IP
                            replacingTextLabels(find, "#IP", sIp),
                            // Замены Габарит
                            replacingTextLabels(find, "#Габарит", sGab),
                            // Замены Марка
                            replacingTextLabels(find, "#Марка", sMark),
                            // Замены Номер
                            replacingTextLabels(find, "#Номер", sNum),
                            // Замены Климат
                            //verifyIsNotFind.Add(
                            //replacingTextLabels(find, "#Климат", sClimate));
                            // Замены Заземление
                            replacingTextLabels(find, "#Заземление", sGround),
                            // Замены Название
                            replacingTextLabels(find, "#Название", sName),
                            // Замена Вставка
                            replacingTextLabels(find, "#Вставка", sPaste),
                            // Замены Заполнение
                            replacingTextLabels(find, "#Заполнение", firstWords),
                            // Замена Склонение
                            replacingTextLabels(find, "#Склонение", secondWords),
                            // Замена Исполнения
                            replacingTextLabels(find, "#Исполнение", sExecution)
                        };

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
                                replacingTextLabels(find, "#Установка", "монтажной плате"));
                        }
                        else if (sIsp == "ДР")
                        {
                            verifyIsNotFind.Add(
                                replacingTextLabels(find, "#Установка", "din-рейках"));
                        }
                        else
                        {
                            verifyIsNotFind.Add(
                                replacingTextLabels(find, "#Установка", "монтажной плате, din-рейках"));
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

                        document.SaveAs($@"{folderName}\Паспорт {factoryNumber.Replace("/", "_")}.docx");

                        int amountSheet = document.ComputeStatistics(WdStatistic.wdStatisticPages, false);

                        // Вызов логгера
                        WriteLog.Logger(folderName, factoryNumber, amountSheet);

                        document.Close();

                        if (!verifyIsNotFind.All(item => item))
                            MessageWarning(
                                $@"В шаблоне не произошла вставка{Environment.NewLine} одного или нескольких значений.",
                                @"Проблема шаблона");

                        // Работа с элементами формы через событие
                        ProgressStep?.Invoke(++numberStep);

                        firstRow++;
                    }
                    while (endRow > firstRow);
                }
                catch (COMException)
                {
                    MessageError(
                        $@"Проверьте имя проекта внимательно,{Environment.NewLine} экстрасенсы говорят что есть ошибки в заполнеии столбцов.",
                        @"Ошибка надстройки");
                    applicationWord?.Quit();
                }
                catch (RuntimeBinderException)
                {
                    MessageError(
                        $@"Проверьте заполнение всех столбцов,{Environment.NewLine} где-то нехватает данных.",
                        @"Ошибка надстройки");
                    applicationWord?.Quit();
                }
                catch (Exception exception)
                {
                    MessageError(
                        exception.ToString(),
                        @"Ошибка надстройки");
                    applicationWord?.Quit();
                }
                finally
                {
                    // Работа с элементами формы через событие
                    ProgressFinal?.Invoke();
                    applicationWord.Quit();
                }
            }).Start();
        }
        /// <summary>
        /// Функция склонения слов
        /// </summary>
        /// <param name="mReplase"></param>
        /// <returns></returns>
        private string FuncReplace(string mReplase)                          
        {
            string[] subs = mReplase.Split(' ');

            var replace = new Dictionary<string, string>()
            {
                { "Щиток", "Щитка"},
                { "Щит", "Щита"},
                { "Шкаф", "Шкафа"},
                { "Устройство", "Устройства"},
                { "Корпус", "Корпуса"},
                { "Ящик", "Ящика"},
                { "Бокс", "Бокса"},
                { "Панель", "Панели"},
                { "распределительный", "распределительного"}
            };

            for (int i = 0; i < subs.Length; i++)
            {
                if (replace.ContainsKey(subs[i]))
                {
                    subs[i] = replace[subs[i]];
                }
            }
            return String.Join(" ", subs);           
        }
    }
}
