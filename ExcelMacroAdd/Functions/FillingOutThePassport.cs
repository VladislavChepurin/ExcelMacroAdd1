using ExcelMacroAdd.Forms;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services;
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
            ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing, value, ref _replaceTypeObj, ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing);

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
                MessageWarning(Properties.Resources.NotJornal, Properties.Resources.NameWorkbook);
                return;
            }

            var firstRow = Cell.Row; // Вычисляем верхний элемент
            var countRow = Cell.Rows.Count; // Вычисляем кол-во выделенных строк
            var endRow = firstRow + countRow - 1;
            var currentRow = firstRow;
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
                Word.Document document = null;
                string numberProject, serialNumber, model, titleLVSwitchgear, paste, designationOfLVSwitchgear, technicalSpecifications, voltage, current, iPRating, cabinetDimensions, secondWords, apparatusMounting, earthingSystem, cabinetMaterialType, mountingType;                      
                dynamic manufacturingData;

                // Цикл переборки строк
                do
                {                    
                    try
                    {
                        string filename;

                        switch (Worksheet.Cells[currentRow, MountingTypeColumn].Value2.ToString())
                        {
                            case "навесное":
                                filename = Path.Combine(pPath, resources.TemplateWall);
                                break;

                            case "встраиваемое":
                                filename = Path.Combine(pPath, resources.TemplateWall);
                                break;

                            case "напольное":
                                filename = Path.Combine(pPath, resources.TemplateFloor);
                                break;

                            case "навесное для IT оборудования":
                                filename = Path.Combine(pPath, resources.TemplateWallIt);
                                break;

                            case "напольное для IT оборудования":
                                filename = Path.Combine(pPath, resources.TemplateFloorIt);
                                break;

                            default:
                                filename = Path.Combine(pPath, resources.TemplateFloor);
                                break;
                        }
                        numberProject = Worksheet.Cells[currentRow, NumberProjectColumn].Value2.ToString().Replace('/', '_');
                        model = Worksheet.Cells[currentRow, ModelColumn].Value2.ToString();
                        titleLVSwitchgear = Worksheet.Cells[currentRow, TitleLVSwitchgearColumn].Value2.ToString();
                        paste = FuncReplace(titleLVSwitchgear ?? string.Empty); // ссылка на метод замены
                        designationOfLVSwitchgear = Worksheet.Cells[currentRow, DesignationOfLVSwitchgearColumn].Value2.ToString();
                        technicalSpecifications = Worksheet.Cells[currentRow, TechnicalSpecificationsColumn].Value2.ToString();
                        voltage = Worksheet.Cells[currentRow, VoltageColumn].Value2.ToString();
                        current = Worksheet.Cells[currentRow, CurrentColumn].Value2.ToString();
                        iPRating = Worksheet.Cells[currentRow, IPRatingColumn].Value2.ToString();                        
                        cabinetDimensions = Worksheet.Cells[currentRow, EnclosureHeightColumn].Value2.ToString() + "x"
                                          + Worksheet.Cells[currentRow, EnclosureWidthColumn].Value2.ToString() + "x"
                                          + Worksheet.Cells[currentRow, EnclosureDepthColumn].Value2.ToString();
                        secondWords = FuncReplace(designationOfLVSwitchgear ?? string.Empty); // ссылка на метод замены
                        manufacturingData = Worksheet.Cells[currentRow, ManufacturingDataColumn].Value2;
                        serialNumber = Worksheet.Cells[currentRow, SerialNumberColumn].Value2.ToString();
                        apparatusMounting = Worksheet.Cells[currentRow, ApparatusMountingColumn].Value2.ToString();
                        earthingSystem = Worksheet.Cells[currentRow, EarthingSystemColumn].Value2.ToString();
                        cabinetMaterialType = Worksheet.Cells[currentRow, CabinetMaterialTypeColumn].Value2.ToString();
                        mountingType = Worksheet.Cells[currentRow, MountingTypeColumn].Value2.ToString();

                        //Открываем Word                   
                        document = applicationWord.Documents.Open(filename, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate,
                            revert, writePasswordDocument, writePasswordTemplate, format, encoding, oVisible, openAndRepair, documentDirection,
                            noEncodingDialog, xmlTransform);
                        applicationWord.Visible = false;
                        //Инициализация метода Find
                        Find find = applicationWord.Selection.Find;                

                        var verifyIsNotFind = new List<bool>(16)
                        {
                            // Замены ТУ
                            replacingTextLabels(find, "#ТУ", technicalSpecifications),
                            // Замены Ток
                            replacingTextLabels(find, "#Ток", current),
                            // Замены IP
                            replacingTextLabels(find, "#IP", iPRating),
                            // Замены Габарит
                            replacingTextLabels(find, "#Габарит", cabinetDimensions),
                            // Замены Марка
                            replacingTextLabels(find, "#Марка", model),
                            // Замены Номер
                            replacingTextLabels(find, "#Номер", serialNumber),
                            // Замены Климат
                            //verifyIsNotFind.Add(
                            //replacingTextLabels(find, "#Климат", sClimate));
                            // Замены Заземление
                            replacingTextLabels(find, "#Заземление", earthingSystem),
                            // Замены Название
                            replacingTextLabels(find, "#Название", titleLVSwitchgear),
                            // Замена Вставка
                            replacingTextLabels(find, "#Вставка", paste),
                            // Замены Заполнение
                            replacingTextLabels(find, "#Заполнение", designationOfLVSwitchgear),
                            // Замена Склонение
                            replacingTextLabels(find, "#Склонение", secondWords),
                            // Замена Исполнения
                            replacingTextLabels(find, "#Исполнение", mountingType),
                            // Замена Даты
                            replacingTextLabels(find, "#Дата", FormingData(manufacturingData)),
                            // Замены Исполнение
                            replacingTextLabels(find, "#Установка", FormingInstalling(apparatusMounting)),
                            // Замены материал корпуса
                            replacingTextLabels(find, "#Корпус", FormingMaterial(cabinetMaterialType)),
                            //Замены напряжения
                            replacingTextLabels(find, "#Напряжение", FormingVoltage(voltage))
                        };

                        //Путь к папке Рабочего стола                                     
                        string folderName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Паспорта " + numberProject);
                        DirectoryInfo drInfo = new DirectoryInfo(folderName);
                        // Проверяем есть ли папка, если нет создаем
                        if (!drInfo.Exists)
                        {
                            drInfo.Create();                          
                        }

                        document.SaveAs($@"{folderName}\Паспорт {serialNumber.Replace("/", "_")}.docx");

                        int amountSheet = document.ComputeStatistics(WdStatistic.wdStatisticPages, false);                                              

                        if (verifyIsNotFind.All(item => item))
                        {
                            // Вызов местного логгера                       
                            ThisWriteLog(folderName, $"{DateTime.Now} | Паспорт {serialNumber} сформирован успешно." +
                                $" В паспорте {amountSheet} листов");
                        }
                        else
                        {
                            // Вызов местного логгера                       
                            ThisWriteLog(folderName, $"{DateTime.Now} | Паспорт {serialNumber} сформирован, но не произошла вставка одного или нескольких значений." +
                                $"  В паспорте {amountSheet} листов");
                        }                                                                     
                    }

                    catch (COMException ex)
                    {
                        MessageError(
                            $@"Проверьте имя проекта внимательно,{Environment.NewLine} экстрасенсы говорят что есть ошибки в заполнеии столбцов.",
                            @"Ошибка надстройки");
                        Logger.LogException(ex);                    
                    }

                    catch (RuntimeBinderException ex)
                    {
                        MessageError(
                            $@"Проверьте заполнение всех столбцов паспорта {null},{Environment.NewLine} где-то нехватает данных.",
                            @"Ошибка надстройки");
                        Logger.LogException(ex);                       
                    }

                    catch (Exception ex)
                    {
                        MessageError(
                            $"Произошла неизветсная ошибка при заполнении паспорта",
                            @"Ошибка надстройки");
                        Logger.LogException(ex);                      
                    }

                    finally
                    {
                        if (document != null)
                        {
                            document.Close();
                            Marshal.ReleaseComObject(document);
                        }

                        // Работа с элементами формы через событие
                        ProgressStep?.Invoke(++numberStep);
                        currentRow++;
                    }                
                }
                while (currentRow <= endRow);

                if (applicationWord != null)
                {
                    applicationWord.Quit();
                    Marshal.ReleaseComObject(applicationWord);
                }

                //Сборка мусора (опционально, но рекомендуется)
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // Работа с элементами UI через событие                  
                ProgressFinal?.Invoke();

            }).Start();
        }

        /// <summary>
        /// Функция форматирует напряжение для паспорта
        /// </summary>
        /// <param name="voltage"></param>
        /// <returns></returns>
        private string FormingVoltage(string voltage)
        {
            switch (voltage)
            {
                case "380":
                    return "~230/380 В.";
                case "220":
                    return "~230В.";
                case "230":
                    return "~230В.";
                default:
                    return voltage + "В.";
            }
        }

        /// <summary>
        /// Функция форматирует материал шкафа в паспорте
        /// </summary>
        /// <param name="material"></param>
        /// <returns></returns>
        private string FormingMaterial(string material)
        {
            switch (material)
            {
                case "Металл":
                    return "металлическом";
                case "Пластик":
                    return "пластиковом";
                case "Композит":
                    return "композитном";
                default:
                    return "металлическом или пластиковом";
            }
        }

        /// <summary>
        /// Функция форматирует установку аппаратов паспорте
        /// </summary>
        /// <param name="installing"></param>
        /// <returns></returns>
        private string FormingInstalling(string installing)
        {
            switch (installing)
            {
                case "МП":
                    return "монтажной плате";
                case "ДР":
                    return "din-рейках";
                case "19\"":
                    return "19\u02EE стойках";
                default:
                    return "монтажной плате, din-рейках";
            }
        }

        /// <summary>
        /// Функция преобразования даты получаемой от Excel
        /// </summary>
        /// <param name="buildData"></param>
        /// <returns></returns>
        private string FormingData(dynamic buildData)
        {
            try
            {
                if (buildData != null & double.TryParse(buildData.ToString(), out double dateValue))
                    return DateTime.FromOADate(dateValue).ToString("D");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
            }
            return "«____» __________ 202_ г.";
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
                { "устройство", "устройства"},
                { "Корпус", "Корпуса"},
                { "Ящик", "Ящика"},
                { "Бокс", "Бокса"},
                { "Панель", "Панели"},
                { "распределительный", "распределительного"},
                { "телекоммуникационный", "телекоммуникационного"},
                { "Источник", "Источника"},
                { "источник", "источника"},
                { "Система", "Системы"},
                { "система", "системы"},
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

        /// <summary>
        /// Метод для записи логов формиррования паспортов
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="saveNum"></param> 
        /// <param name="amount"></param>        
        private void ThisWriteLog(string folder, string logText)
        {
            string patch = Path.Combine(folder, "log.txt");
            using (StreamWriter output = File.AppendText(patch))
            {
                output.WriteLine(logText);
            }
        }
    }
}
