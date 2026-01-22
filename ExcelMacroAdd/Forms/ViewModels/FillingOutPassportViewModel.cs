using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelMacroAdd.Forms.ViewModels
{
    internal class FillingOutPassportViewModel : AbstractFunctions, INotifyPropertyChanged
    {       
        // Переменные иницализации                   
        private readonly object _confirmConversions = false;
        private readonly object _readOnly = false;
        private readonly object _addToRecentFiles = false;
        private readonly object _passwordDocument = Type.Missing;
        private readonly object _passwordTemplate = Type.Missing;
        private readonly object _revert = false;
        private readonly object _writePasswordDocument = Type.Missing;
        private readonly object _writePasswordTemplate = Type.Missing;
        private readonly object _format = Type.Missing;
        private readonly object _encoding = Type.Missing;
        private readonly object _oVisible = Type.Missing;
        private readonly object _openAndRepair = Type.Missing;
        private readonly object _documentDirection = Type.Missing;
        private readonly object _noEncodingDialog = false;
        private readonly object _xmlTransform = Type.Missing;
        private static object _replaceTypeObj = WdReplace.wdReplaceAll;
        private static object _wordMissing = Missing.Value;
        private readonly IFillingOutThePassportSettings _resources;
        private readonly string _pPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template");

        private readonly SynchronizationContext _syncContext;

        readonly Func<Find, string, string, bool> replacingTextLabels = (find, label, value) => find.Execute(label, ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing,
            ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing, value, ref _replaceTypeObj, ref _wordMissing, ref _wordMissing, ref _wordMissing, ref _wordMissing);

        #region Binding Property
        private string _infoLabelText;
        private bool _isEnabledBtnClose;
        private int _progressBarMinimum;
        private int _progressBarMaximum;
        private int _progressBarStep;
        private int _progressBarValue;

        public bool IsEnabledBtnClose
        {
            get => _isEnabledBtnClose;
            set { _isEnabledBtnClose = value; OnPropertyChanged(nameof(IsEnabledBtnClose)); }
        }

        public string InfoLabelText
        {
            get => _infoLabelText;
            set { _infoLabelText = value; OnPropertyChanged(nameof(InfoLabelText)); }
        }

        public int ProgressBarMinimum
        {
            get => _progressBarMinimum;
            set { _progressBarMinimum = value; OnPropertyChanged(nameof(ProgressBarMinimum)); }
        }

        public int ProgressBarMaximum
        {
            get => _progressBarMaximum;
            set { _progressBarMaximum = value; OnPropertyChanged(nameof(ProgressBarMaximum)); }
        }

        public int ProgressBarStep
        {
            get => _progressBarStep;
            set { _progressBarStep = value; OnPropertyChanged(nameof(ProgressBarStep)); }
        }

        public int ProgressBarValue
        {
            get => _progressBarValue;
            set { _progressBarValue = value; OnPropertyChanged(nameof(ProgressBarValue)); }
        }

        #endregion


        public FillingOutPassportViewModel(IFillingOutThePassportSettings resources)
        {                        
            // Сохраняем контекст синхронизации UI-потока
            _syncContext = SynchronizationContext.Current;
            this._resources = resources;
        }

        public override void Start()
        {            
            if (Application.ActiveWorkbook.Name != _resources.NameFileJournal) // Проверка по имени книги
            {
                MessageWarning(Properties.Resources.NotJornal, Properties.Resources.NameWorkbook);
                RequestClose?.Invoke(this, EventArgs.Empty);
                return;
            }

            var firstRow = Cell.Row; // Вычисляем верхний элемент
            var countRow = Cell.Rows.Count; // Вычисляем кол-во выделенных строк
            var endRow = firstRow + countRow - 1;
            var currentRow = firstRow;
        

            ProgressBarMinimum = 0;
            ProgressBarStep = 1;
            ProgressBarMaximum = countRow;

            System.Threading.Tasks.Task.Run(() =>
            {
                //Инициализируем параметры Word
                Word.Application applicationWord = new Word.Application();              
                // Переменная объект документа
                Word.Document document = null;
                string numberProject, serialNumber, model, titleLVSwitchgear, paste, designationOfLVSwitchgear, technicalSpecifications, voltage, current, iPRating, cabinetDimensions, secondWords, apparatusMounting, earthingSystem, cabinetMaterialType, mountingType;                      
                dynamic manufacturingData;

                int step = 0;
                // Цикл переборки строк
                do
                {                    
                    try
                    {
                        string filename;
                        string fileSHA1;
                        
                        switch (Worksheet.Cells[currentRow, MountingTypeColumn].Value2.ToString())
                        {
                            case "навесное":
                                filename = _resources.TemplateWall;
                                fileSHA1 = _resources.TemplateWallSHA1;
                                break;

                            case "встраиваемое":
                                filename = _resources.TemplateWall;
                                fileSHA1 = _resources.TemplateWall;
                                break;

                            case "напольное":
                                filename = _resources.TemplateFloor;
                                fileSHA1 = _resources.TemplateFloorSHA1;
                                break;

                            case "навесное для IT оборудования":
                                filename = _resources.TemplateWallIt;
                                fileSHA1 = _resources.TemplateWallItSHA1;
                                break;

                            case "напольное для IT оборудования":
                                filename = _resources.TemplateFloorIt;
                                fileSHA1 = _resources.TemplateFloorItSHA1;
                                break;

                            default:
                                filename = _resources.TemplateFloor;
                                fileSHA1 = _resources.TemplateFloorSHA1;
                                break;
                        }

                        string fileFullPath = Path.Combine(_pPath, filename);                    

                        if (_resources.CheckSHA1)
                        {                         
                            if (TemplateFileSHA1(fileFullPath) != fileSHA1)
                            {
                                DialogResult dialogResult = MessageBox.Show($"Ошибка в контрольной сумме файла {filename}. Нажмите ДА если хотите автоматически восстановить файл.", "Ошибка шаблона", MessageBoxButtons.YesNo);
                                if (dialogResult == DialogResult.Yes)
                                {
                                    string backupFile = String.Concat(fileFullPath, ".bak");

                                    // Проверяем существование файлов
                                    if (!File.Exists(backupFile))
                                    {
                                        MessageError($"Резервная копия {backupFile} не найдена!", "Ошибка файла");
                                        break;                                        
                                    }

                                    // Если оригинальный файл существует, заменяем его
                                    if (File.Exists(fileFullPath))
                                    {
                                        File.Replace(backupFile, fileFullPath, fileFullPath + ".old", true);
                                        Console.WriteLine("Файл успешно восстановлен из резервной копии!");
                                    }
                                    else
                                    {
                                        // Если оригинального файла нет, просто копируем backup
                                        File.Copy(backupFile, fileFullPath);
                                        Console.WriteLine("Файл восстановлен из резервной копии!");
                                    }
                                }
                                break;
                            }
                        }

                        numberProject = Worksheet.Cells[currentRow, NumberProjectColumn].Value2.ToString().Replace('/', '_').Trim();
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
                        document = applicationWord.Documents.Open(fileFullPath, _confirmConversions, _readOnly, _addToRecentFiles, _passwordDocument, _passwordTemplate,
                            _revert, _writePasswordDocument, _writePasswordTemplate, _format, _encoding, _oVisible, _openAndRepair, _documentDirection,
                            _noEncodingDialog, _xmlTransform);
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

                        // Работа с элементами UI
                        _syncContext.Post(_ =>
                        {
                            progressBarUI(++step);                                  
                            InfoLabelText = $@"Подождите пожайлуста, идет заполнение паспортов {step}/{countRow}.";

                        }, null);

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
            })
                .ContinueWith(t =>
                {
                    // Этот код выполнится после завершения фоновой задачи
                    if (t.IsCompleted)
                    {
                        InfoLabelText = "Паспота заполнены.Ты молодец";
                        IsEnabledBtnClose = true;
                    }
                    else if (t.IsFaulted)
                    {
                        // Обработка ошибок
                        Logger.LogException(t.Exception);
                    }
                }, TaskScheduler.FromCurrentSynchronizationContext()); // Гарантирует выполнение в UI потоке
        }

        private static string TemplateFileSHA1(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("Path cannot be null or empty", nameof(path));

            if (!File.Exists(path))
                throw new FileNotFoundException($"File not found: {path}", path);

            try
            {
                using (var stream = File.OpenRead(path))
                {
                    using (var sha1 = System.Security.Cryptography.SHA1.Create())
                    {
                        byte[] hash = sha1.ComputeHash(stream);
                        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to compute SHA1 for file: {path}", ex);
            }
        }


        private void progressBarUI(int step)
        {
            if (step == ProgressBarMaximum)
            {
                // Special case as value can't be set greater than Maximum.
                ProgressBarMaximum = step + 1;     // Temporarily Increase Maximum
                ProgressBarValue = step;           // Move past
                ProgressBarMaximum = step;         // Reset maximum
            }
            else
            {
                ProgressBarValue = step + 1;       // Move past
            }
            //ProgressBarValue = step;             // Move to correct value
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
                { "Термошкаф", "Термошкафа" },
                { "термошкаф", "термошкафа" }
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

        public event EventHandler RequestClose;

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
