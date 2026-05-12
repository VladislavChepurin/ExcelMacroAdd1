using ExcelMacroAdd.BusinessLayer;
using ExcelMacroAdd.Forms;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.ProxyObjects;
using ExcelMacroAdd.Serializable;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services;
using ExcelMacroAdd.Services.Interfaces;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Infrastructure;
using ExcelMacroAdd.DataLayer.UnitOfWork;
using Office = Microsoft.Office.Core;



namespace ExcelMacroAdd
{
    [ComVisible(true)]
    public class NewRibbon : Office.IRibbonExtensibility, IDisposable
    {
        private static readonly TimeSpan GlobalDatabaseProbeTimeout = TimeSpan.FromSeconds(2);
        private Office.IRibbonUI ribbon;
        private readonly string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config/appSettings.json");
        private readonly IDataInXml dataInXml;
        private readonly IFillingOutThePassportSettings resources;
        private readonly ICorrectFontResources correctFontResources;
        private readonly IFormSettings formSettings;
        private readonly ITypeNkySettings[] typeNkySettings;
        private readonly IAdditionalDevicesService additionalDevicesService;
        private readonly ICircuitBreakerService circuitBreakerService;
        private readonly IJournalNkuService journalNkuService;
        private readonly IJournalNkuWriteService journalNkuWriteService;
        private readonly INotPriceComponentsService notPriceComponentsService;
        private readonly ISwitchService switchService;
        private readonly ITransformerService transformerService;
        private readonly ITwinBlockService twinBlockService;
        private readonly IUnitOfWorkFactory unitOfWorkFactory;
        private readonly bool locationDataBase = default;
        private readonly IMemoryCache memoryCache;
        private readonly IValidateLicenseKey validateLicenseKey;
        private bool _disposed;
        private volatile bool _notPriceComponentsOpen = false;
        private volatile bool _termoCalculationOpen = false;
        private volatile bool _selectionTransformer = false;
        private volatile bool _selectionTwinBlock = false;
        private volatile bool _selectionModularDevices = false;

        public NewRibbon()
        {
            AppSettingsDeserialize app = new AppSettingsDeserialize(jsonFilePath);
            var settings = app.GetSettingsModels();
            resources = settings.Resources;
            correctFontResources = settings.CorrectFontResources;
            formSettings = settings.FormSettings;
            typeNkySettings = settings.TypeNkySettings;
            var cacheOptions = new MemoryCacheOptions
            {
                ExpirationScanFrequency = TimeSpan.FromMinutes(30)
            };
            memoryCache = new MemoryCache(cacheOptions);

            string path;
            path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DataLayer/DataBase/");

            if (settings.GlobalDateBaseLocationEnable &&
                TryProbeGlobalDatabase(settings.GlobalDateBaseLocation, GlobalDatabaseProbeTimeout))
            {
                path = settings.GlobalDateBaseLocation;
                locationDataBase = true;
            }

            var appContextFactory = new AppContextFactory(path);
            unitOfWorkFactory = new UnitOfWorkFactory(appContextFactory);

            additionalDevicesService = new AdditionalDevicesQueryService(unitOfWorkFactory);
            circuitBreakerService = new CircuitBreakerQueryService(unitOfWorkFactory, memoryCache);
            journalNkuService = new JournalNkuQueryService(unitOfWorkFactory, memoryCache);
            journalNkuWriteService = new JournalNkuWriteService(unitOfWorkFactory, memoryCache);
            notPriceComponentsService = new NotPriceComponentsService(unitOfWorkFactory);
            switchService = new SwitchQueryService(unitOfWorkFactory, memoryCache);
            transformerService = new TransformerQueryService(unitOfWorkFactory);
            twinBlockService = new TwinBlockQueryService(unitOfWorkFactory);
            validateLicenseKey = new ValidateLicenseKey();

            //Создание внедряемых зависимостей
            dataInXml = new DataInXmlProxy(new DataInXml());

#if !DEBUG
            //Чтобы не тормозил интерфейс при первом запросе в базу данных
            new Task(() =>
            {
                if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DataLayer/DataBase/BdMacro.sqlite")))
                {
                    using (var warmupUnitOfWork = unitOfWorkFactory.Create())
                    {
                        warmupUnitOfWork.Context.Switches.Select(x => x.Id).FirstOrDefault();
                    }
                }
            }).Start();
#endif
        }

        #region Элементы IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelMacroAdd.NewRibbon.xml");
        }

        #endregion

        #region Обратные вызовы ленты
        //Информацию о методах создания обратного вызова см. здесь. Дополнительные сведения о методах добавления обратного вызова см. по ссылке https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public Image GetImage(string ImageName)
        {
            return (Image)Properties.Resources.ResourceManager.GetObject(ImageName);
        }

        public async Task OnActionCallbackBase(Office.IRibbonControl control)
        {
//#if !DEBUG
            if (!validateLicenseKey.ValidateKey())
            {

                MessageBox.Show(Properties.Resources.LicenseText, "Внимание");
                return;
            }
//#endif
            switch (control.Id)
            {
                //Заполнение паспортов
                case "FillingOutThePassport_Button":
                    var fillingOutThePassport = new FillingOutPassports(resources);
                    fillingOutThePassport.Show();
                    break;

                //Удалить формулы выделенной области
                case "DeleteFormula_Button":
                    var deleteFormula = new DeleteFormula();
                    deleteFormula.Start();
                    break;

                //Удалить все формулы
                case "DeleteAllFormula_Button":
                    var deleteAllFormula = new DeleteAllFormula();
                    deleteAllFormula.Start();
                    break;

                //Корпуса щитов
                case "BoxShield_Button":
                    if (journalNkuService != null)
                    {
                        var boxShield = new BoxShield(journalNkuService, resources);
                        await boxShield.StartAsync();
                    }
                    break;

                //Корпуса в базу
                case "AddBoxDb_Button":
                    if (journalNkuWriteService != null)
                    {
                        var addBoxDb = new AddBoxDb(journalNkuWriteService, resources);
                        await addBoxDb.StartAsync();
                    }
                    break;

                //Исправить запись в БД
                case "CorrectDb_Button":
                    if (journalNkuWriteService != null)
                    {
                        var correctDb = new CorrectDb(journalNkuWriteService, resources);
                        await correctDb.StartAsync();
                    }
                    break;
            }
        }

        public void OnActionCallbackDecoration(Office.IRibbonControl control)
        {
//#if !DEBUG
            if (!validateLicenseKey.ValidateKey())
            {
                MessageBox.Show(Properties.Resources.LicenseText, "Внимание");
                return;
            }
//#endif
            switch (control.Id)
            {
                //Разметка листов
                case "Linker_Button":
                    var linker = new Linker(correctFontResources);
                    linker.Start();
                    break;

                //Границы таблицы
                case "BordersTable_Button":
                    var bordersTable = new BordersTable();
                    bordersTable.Start();
                    break;

                //Правка шрифта
                case "CorrectFont_Button":
                    var correctFont = new CorrectFont(correctFontResources);
                    correctFont.Start();
                    break;

                // Разметка таблицы расчетов
                case "CalculationMarkup_Button":
                    var calculationMarkup = new CalculationMarkup(correctFontResources);
                    calculationMarkup.Start();
                    break;

                // Исправление расчетов
                case "EditCalculation_Button":
                    var editCalculation = new EditCalculation(correctFontResources);
                    editCalculation.Start();
                    break;

                // Объединение ячеек
                case "CombiningCells_Button":
                    var combiningCells = new CombiningCells();
                    combiningCells.Start();
                    break;
            }
        }

        public void OnActionCallbackSearch(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                //Поиск в Яндексе
                case "Yandex_Button":
                    var yandexSearch = new InternetSearch("http://www.yandex.ru/yandsearch?text=");
                    yandexSearch.Start();
                    break;

                //Поиск в Google
                case "Google_Button":
                    var googleSearch = new InternetSearch("https://www.google.ru/search?q=");
                    googleSearch.Start();
                    break;
            }
        }

        public async Task OnActionCallbackCalculation(Office.IRibbonControl control)
        {
//#if !DEBUG
            if (!validateLicenseKey.ValidateKey())
            {
                MessageBox.Show(Properties.Resources.LicenseText, "Внимание");
                return;
            }
//#endif
            switch (control.Id)
            {
                //Вставка формулы Iek
                case "Iek_Button":
                    RunWriteExcel("IEK");
                    break;

                //Вставка формулы Ekf
                case "Ekf_Button":
                    RunWriteExcel("EKF");
                    break;

                //Вставка формулы Dkc
                case "Dkc_Button":
                    RunWriteExcel("DKC");
                    break;

                //Вставка формулы Keaz
                case "Keaz_Button":
                    RunWriteExcel("KEAZ");
                    break;

                //Вставка формулы Dek
                case "Dek_Button":
                    RunWriteExcel("DEKraft");
                    break;

                //Вставка формулы Chint
                case "Chint_Button":
                    RunWriteExcel("Chint");
                    break;

                //Модульные аппараты
                case "SelectionModularDevices_Button":
                    if (circuitBreakerService == null || switchService == null || additionalDevicesService == null) break;
                    if (_selectionModularDevices)
                    {
                        MessageBox.Show("Окно уже открыто", "Информация",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                    _selectionModularDevices = true;
                    try
                    {
                        await ShowFormOnStaThread(() => new SelectionModularDevices(
                            dataInXml,
                            circuitBreakerService,
                            switchService,
                            additionalDevicesService,
                            formSettings));
                    }
                    finally
                    {
                        _selectionModularDevices = false;
                    }
                    break;

                //Трансформаторы тока
                case "SelectionTransformer_Button":
                    if (transformerService == null) break;
                    if (_selectionTransformer)
                    {
                        MessageBox.Show("Окно уже открыто", "Информация",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                    _selectionTransformer = true;
                    try
                    {
                        await ShowFormOnStaThread(() => new SelectionTransformer(dataInXml, transformerService, formSettings));
                    }
                    finally
                    {
                        _selectionTransformer = false;
                    }
                    break;

                //Рубильники TwinBlock
                case "SelectionTwinBlock_Button":
                    if (twinBlockService == null) break;
                    if (_selectionTwinBlock)
                    {
                        MessageBox.Show("Окно уже открыто", "Информация",
                          MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }

                    _selectionTwinBlock = true;
                    try
                    {
                        await ShowFormOnStaThread(() => new SelectionTwinBlock(dataInXml, twinBlockService, formSettings));
                    }
                    finally
                    {
                        _selectionTwinBlock = false;
                    }
                    break;

                //Расчет обогрева
                case "TermoCalculation_Button":
                    if (journalNkuService == null) break;
                    if (_termoCalculationOpen)
                    {
                        MessageBox.Show("Окно уже открыто", "Информация",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                    _termoCalculationOpen = true;
                    try
                    {
                        await ShowFormOnStaThread(() => new TermoCalculation(journalNkuService, formSettings));
                    }
                    finally
                    {
                        _termoCalculationOpen = false;
                    }
                    break;


                //Не тарифные позиции
                case "NotPriceComponent_Button":
                    if (notPriceComponentsService == null) break;
                    if (_notPriceComponentsOpen)
                    {
                        MessageBox.Show("Окно уже открыто", "Информация",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                    _notPriceComponentsOpen = true;
                    try
                    {
                        await ShowFormOnStaThread(() => new NotPriceComponents(notPriceComponentsService, formSettings));
                    }
                    finally
                    {
                        _notPriceComponentsOpen = false;
                    }
                    break;

                //Таблица типов
                case "TypeNky_Button":
                    // Проверяем, есть ли уже такая панель
                    var existingPane = Globals.ThisAddIn.CustomTaskPanes
                        .FirstOrDefault(p => p.Title == "Тип шкафов");

                    if (existingPane == null)
                    {
                        var typeNky = new TypeNky(typeNkySettings);
                        existingPane = Globals.ThisAddIn.CustomTaskPanes.Add(typeNky, "Тип шкафов");
                        existingPane.Width = 400;
                        existingPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    }
                    existingPane.Visible = !existingPane.Visible;

                    break;
            }
        }

        /// <summary>
        /// Обёртка для вставки формул ВПР через кнопки Ribbon.
        /// Внешний ExcelPerformanceScope гарантирует:
        ///   - ScreenUpdating, Calculation, Events отключены на время вставки
        ///   - Calculate() вызывается ОДИН раз в конце
        ///   - Если лист уже пересчитывался в этой сессии — Calculate() пропускается
        /// При повторных нажатиях кнопок ВПР на том же листе
        /// формулы вставляются мгновенно без пересчёта.
        /// </summary>
        private void RunWriteExcel(string vendor)
        {
            using (var scope = new ExcelPerformanceScope(Globals.ThisAddIn.GetApplication()))
            {
                var writeExcel = new WriteExcel(dataInXml, vendor);
                writeExcel.Start();
            }
        }

        public async Task OnActionCallbackOther(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                //Окно настроек
                case "Settings_Button":
                    await ShowFormOnStaThread(() => new Settings(dataInXml));
                    break;

                //Окно о программе
                case "About_Button":
                    await ShowFormOnStaThread(() => new AboutBox1(locationDataBase));
                    break;

                //Открыть папку
                case "Open_Button":
                    Process.Start("explorer.exe", AppDomain.CurrentDomain.BaseDirectory);
                    break;
            }
        }

        //public string GetLabelText(Office.IRibbonControl control)
        //{
        //    return locationDataBase ? Properties.Resources.Global : Properties.Resources.Local;         
        //}

        #endregion

        #region Вспомогательные методы

        /// <summary>
        /// Освобождает ресурсы аддина.
        /// Вызывается из ThisAddIn.Shutdown при выгрузке аддина.
        /// </summary>
        public void Dispose()
        {
            if (!_disposed)
            {
                (memoryCache as IDisposable)?.Dispose();
                _disposed = true;
            }
        }

        /// <summary>
        /// Создаёт отдельный STA-поток с message pump для показа WinForms-диалога.
        /// 
        /// Почему нельзя Task.Run:
        ///   Task.Run использует потоки из ThreadPool, которые работают в MTA-режиме.
        ///   WinForms требует STA-поток для корректной работы message loop, 
        ///   ComboBox, DataGridView, Clipboard и других контролов.
        ///   Без STA возможны зависания, мерцание, крэши при Drag&Drop и буфере обмена.
        ///
        /// Как работает:
        ///   1. Создаём Thread с ApartmentState.STA
        ///   2. Внутри потока вызываем Application.Run(form) — это запускает полноценный message loop
        ///   3. TaskCompletionSource позволяет await-ить завершение потока из вызывающего кода
        ///   4. Форма корректно Dispose-ится при закрытии
        /// </summary>
        private static Task ShowFormOnStaThread<T>(Func<T> formFactory) where T : System.Windows.Forms.Form
        {
            var tcs = new TaskCompletionSource<bool>();

            var thread = new Thread(() =>
            {
                try
                {
                    using (var form = formFactory())
                    {
                        System.Windows.Forms.Application.Run(form);
                    }
                    tcs.TrySetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();

            return tcs.Task;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        private static bool TryProbeGlobalDatabase(string globalDatabaseLocation, TimeSpan timeout)
        {
            if (string.IsNullOrWhiteSpace(globalDatabaseLocation))
            {
                return false;
            }

            string databasePath = Path.Combine(globalDatabaseLocation, "BdMain.sqlite");
            var probeTask = Task.Run(() => File.Exists(databasePath));

            if (!probeTask.Wait(timeout))
            {
                return false;
            }

            return probeTask.Status == TaskStatus.RanToCompletion && probeTask.Result;
        }

        #endregion
    }
}
