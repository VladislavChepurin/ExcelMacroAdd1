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
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;
using Office = Microsoft.Office.Core;



namespace ExcelMacroAdd
{
    [ComVisible(true)]
    public class NewRibbon : Office.IRibbonExtensibility, IDisposable
    {
        private Office.IRibbonUI ribbon;
        private readonly string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config/appSettings.json");
        private readonly IDataInXml dataInXml;
        private readonly IFillingOutThePassportSettings resources;
        private readonly ICorrectFontResources correctFontResources;
        private readonly IFormSettings formSettings;
        private readonly ITypeNkySettings[] typeNkySettings;
        private readonly AccessData accessData;
        private readonly bool locationDataBase = default;
        private readonly IMemoryCache memoryCache;
        private readonly IValidateLicenseKey validateLicenseKey;
        private bool _disposed;

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

            if (settings.GlobalDateBaseLocationEnable && File.Exists(settings.GlobalDateBaseLocation + "BdMain.sqlite"))
            {
                path = settings.GlobalDateBaseLocation;
                locationDataBase = true;
            }
            else
            {
                path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DataLayer/DataBase/");
            }

            var context = new AppContext(path);
            accessData = new AccessData(context, memoryCache);
            validateLicenseKey = new ValidateLicenseKey(settings.LineseKey);

            //Создание внедряемых зависимостей
            dataInXml = new DataInXmlProxy(new DataInXml());

#if !DEBUG
            //Чтобы не тормозил интерфейс при первом запросе в базу данных
            new Task(() =>
            {
                if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DataLayer/DataBase/BdMacro.sqlite")))
                {
                    context.Switches.AsParallel().Select(x => x.Id).ToList();
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
#if !DEBUG
            if (!validateLicenseKey.ValidateKey())
            {

                MessageBox.Show(Properties.Resources.LicenseText, "Внимание");
                return;
            }
#endif
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
                    if (accessData != null)
                    {
                        var boxShield = new BoxShield(accessData, resources);
                        await boxShield.StartAsync();
                    }
                    break;

                //Корпуса в базу
                case "AddBoxDb_Button":
                    if (accessData != null)
                    {
                        var addBoxDb = new AddBoxDb(accessData, resources);
                        await addBoxDb.StartAsync();
                    }
                    break;

                //Исправить запись в БД
                case "CorrectDb_Button":
                    if (accessData != null)
                    {
                        var correctDb = new CorrectDb(accessData, resources);
                        await correctDb.StartAsync();
                    }
                    break;
            }
        }

        public void OnActionCallbackDecoration(Office.IRibbonControl control)
        {
#if !DEBUG
            if (!validateLicenseKey.ValidateKey())
            {
                MessageBox.Show(Properties.Resources.LicenseText, "Внимание");
                return;
            }
#endif
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
#if !DEBUG
            if (!validateLicenseKey.ValidateKey())
            {
                MessageBox.Show(Properties.Resources.LicenseText, "Внимание");
                return;
            }
#endif
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
                    if (accessData != null)
                    {
                        await ShowFormOnStaThread(() => new SelectionModularDevices(dataInXml, accessData, formSettings));
                    }

                    break;

                //Трансформаторы тока
                case "SelectionTransformer_Button":
                    if (accessData != null)
                        await ShowFormOnStaThread(() => new SelectionTransformer(dataInXml, accessData, formSettings));
                    break;

                //Рубильники TwinBlock
                case "SelectionTwinBlock_Button":
                    if (accessData != null)
                        await ShowFormOnStaThread(() => new SelectionTwinBlock(dataInXml, accessData, formSettings));

                    break;

                //Расчет обогрева
                case "TermoCalculation_Button":
                    if (accessData != null)
                        await ShowFormOnStaThread(() => new TermoCalculation(accessData, formSettings));

                    break;


                //Не тарифные позиции
                case "NotPriceComponent_Button":
                    if (accessData != null)
                        await ShowFormOnStaThread(() => new NotPriceComponents(accessData, formSettings));

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
        /// Освобождает ресурсы аддина: DbContext (через AccessData) и MemoryCache.
        /// Вызывается из ThisAddIn.Shutdown при выгрузке аддина.
        /// </summary>
        public void Dispose()
        {
            if (!_disposed)
            {
                accessData?.Dispose();
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

        #endregion
    }
}