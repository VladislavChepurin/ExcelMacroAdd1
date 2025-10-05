using ExcelMacroAdd.BisinnesLayer;
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
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;
using Office = Microsoft.Office.Core;



namespace ExcelMacroAdd
{
    [ComVisible(true)]
    public class NewRibbon : Office.IRibbonExtensibility
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

            if  (settings.GlobalDateBaseLocationEnable && File.Exists(settings.GlobalDateBaseLocation + "BdMain.sqlite") )
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

        public void OnActionCallbackBase(Office.IRibbonControl control)
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
                        boxShield.Start();
                    }
                    break;

                //Корпуса в базу
                case "AddBoxDb_Button":
                    if (accessData != null)
                    {
                        var addBoxDb = new AddBoxDb(accessData, resources);
                        addBoxDb.Start();
                    }
                    break;

                //Исправить запись в БД
                case "CorrectDb_Button":
                    if (accessData != null)
                    {
                        var correctDb = new CorrectDb(accessData, resources);
                        correctDb.Start();
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
            WriteExcel writeExcel;

            switch (control.Id)
            {
                //Вставка формулы Iek
                case "Iek_Button":
                    writeExcel = new WriteExcel(dataInXml, "IEK");
                    writeExcel.Start();
                    break;

                //Вставка формулы Ekf
                case "Ekf_Button":                  
                    writeExcel = new WriteExcel(dataInXml, "EKF");
                    writeExcel.Start();
                    break;

                //Вставка формулы Dkc
                case "Dkc_Button":                  
                    writeExcel = new WriteExcel(dataInXml, "DKC");
                    writeExcel.Start();
                    break;

                //Вставка формулы Keaz
                case "Keaz_Button":                   
                    writeExcel = new WriteExcel(dataInXml, "KEAZ");
                    writeExcel.Start();
                    break;

                //Вставка формулы Dek
                case "Dek_Button":
                    //writeExcel = new WriteExcel(dataInXml, "Dekraft");
                    writeExcel = new WriteExcel(dataInXml, "DEKraft");
                    writeExcel.Start();
                    break;

                //Вставка формулы Chint
                case "Chint_Button":
                    writeExcel = new WriteExcel(dataInXml, "Chint");
                    writeExcel.Start();
                    break;

                //Модульные аппараты
                case "SelectionModularDevices_Button":
                    if (accessData != null) {
                        await Task.Run(() =>
                        {
                            var selectionModularDevices = new SelectionModularDevices(dataInXml, accessData, formSettings);
                            selectionModularDevices.ShowDialog();
                        });     
                    }

                    break;

                //Трансформаторы тока
                case "SelectionTransformer_Button":
                    if (accessData != null)
                        await Task.Run(() =>
                        {
                            var selectionTransformer = new SelectionTransformer(dataInXml, accessData, formSettings);
                            selectionTransformer.ShowDialog();
                        });
                    break;

                //Рубильники TwinBlock
                case "SelectionTwinBlock_Button":
                    if (accessData != null)
                        await Task.Run(() =>
                        {
                            var selectionTwinBlock = new SelectionTwinBlock(dataInXml, accessData, formSettings);
                            selectionTwinBlock.ShowDialog();
                        });

                    break;

                //Расчет обогрева
                case "TermoCalculation_Button":
                    if (accessData != null)
                        await Task.Run(() =>
                        {
                            var termoCalculation = new TermoCalculation(accessData, formSettings);
                            termoCalculation.ShowDialog();
                        });

                    break;


                //Расчет обогрева
                case "NotPriceComponent_Button":
                    if (accessData != null)
                        await Task.Run(() =>
                        {
                            var notPriceComponents = new NotPriceComponents(accessData, formSettings);
                            notPriceComponents.ShowDialog();
                        });

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
                    existingPane.Visible = true;
                
                    break;

                //Таблица типов
                case "AutoCadAdd":
                    var autoCadCalled = new AutoCadCalled();
                    autoCadCalled.Start();
                    break;
            }
        }

        public async Task OnActionCallbackOther(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                //Окно о программе
                case "Settings_Button":
                    await Task.Run(() =>
                    {
                        Settings fs = new Settings(dataInXml);
                        fs.ShowDialog();
                        Thread.Sleep(5000);
                    });
                    break;

                //Окно о программе
                case "About_Button":
                    await Task.Run(() =>
                    {
                        var about = new AboutBox1(locationDataBase);
                        about.ShowDialog();
                        Thread.Sleep(5000);
                    });
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
