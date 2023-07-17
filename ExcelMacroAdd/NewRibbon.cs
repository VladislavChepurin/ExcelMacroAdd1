using ExcelMacroAdd.BisinnesLayer;
using ExcelMacroAdd.Forms;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.ProxyObjects;
using ExcelMacroAdd.Serializable;
using ExcelMacroAdd.Services;
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
    public class NewRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private readonly string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config/appSettings.json");
        private readonly IDataInXml dataInXml;
        private readonly IResources resources;
        private readonly ICorrectFontResources correctFontResources;
        private readonly IFormSettings formSettings;
        private readonly AccessData accessData;
        private readonly bool locationDataBase = default;

        public NewRibbon()
        {
            AppSettingsDeserialize app = new AppSettingsDeserialize(jsonFilePath);
            var settings = app.GetSettingsModels();
            resources = settings.Resources;
            correctFontResources = settings.CorrectFontResources;
            formSettings = settings.FormSettings;
                       
            string path;
            if (File.Exists(settings.GlobalDateBaseLocation + "BdMacro.sqlite"))
            {
                path = settings.GlobalDateBaseLocation;
                locationDataBase = true;
            }
            else
            {
                path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DataLayer/DataBase/");       
            }
            //Создание внедряемых зависимостей
            dataInXml = new DataInXmlProxy(new Lazy<DataInXml>());
            var context = new AppContext(path);
            accessData = new AccessData(context);
            //Будет утекать 50МБ памяти
            new Thread(() =>
            {
                if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DataLayer/DataBase/BdMacro.sqlite")))
                {
                    context.Switchs.Select(x => x.Id).ToList();
                }
            }).Start();
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
            switch (control.Id)
            {
                //Заполнение паспортов
                case "FillingOutThePassport_Button":
                    var fillingOutThePassport = new FillingOutThePassport(resources);
                    fillingOutThePassport.Start();
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
                    var boxShield = new BoxShield(accessData, resources);
                    boxShield.Start();
                    break;

                //Корпуса в базу
                case "AddBoxDb_Button":
                    var addBoxDb = new AddBoxDb(accessData, resources);
                    addBoxDb.Start();
                    break;
                //Исправить запись в БД
                case "CorrectDb_Button":
                    var correctDb = new CorrectDb(accessData, resources);
                    correctDb.Start();
                    break;
            }
        }

        public void OnActionCallbackDecoration(Office.IRibbonControl control)
        {
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
            }
        }

        public void OnActionCallbackSearch(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                //Поиск в Яндексе
                case "Yandex_Button":
                    var yandexSearch = new YandexSearch();
                    yandexSearch.Start();
                    break;
                //Поиск в Google
                case "Google_Button":
                    var googleSearch = new GoogleSearch();  
                    googleSearch.Start();
                    break;
            }
        }

        public async Task OnActionCallbackCalculation(Office.IRibbonControl control)
        {
            WriteExcel writeExcel;

            switch (control.Id)
            {
                //Вставка формулы Iek
                case "Iek_Button":
                    writeExcel = new WriteExcel(dataInXml, "Iek");
                    writeExcel.Start();
                    break;
                //Вставка формулы Ekf
                case "Ekf_Button":
                    writeExcel = new WriteExcel(dataInXml, "Ekf");
                    writeExcel.Start();
                    break;
                //Вставка формулы Dkc
                case "Dkc_Button":
                    writeExcel = new WriteExcel(dataInXml, "Dkc");
                    writeExcel.Start();
                    break;
                //Вставка формулы Keaz
                case "Keaz_Button":
                    writeExcel = new WriteExcel(dataInXml, "Keaz");
                    writeExcel.Start();
                    break;
                //Вставка формулы Dek
                case "Dek_Button":
                    writeExcel = new WriteExcel(dataInXml, "Dekraft");
                    writeExcel.Start();
                    break;
                //Вставка формулы Chint
                case "Chint_Button":
                    writeExcel = new WriteExcel(dataInXml, "Chint");
                    writeExcel.Start();
                    break;
                //Модульные аппараты
                case "SelectionCircuitBreaker_Button":
                    await SelectionCircuitBreaker.getInstance(dataInXml, accessData, formSettings);
                    break;
                //Трансформаторы тока
                case "SelectionTransformer_Button":
                    await SelectionTransformer.getInstance(dataInXml, accessData, formSettings);
                    break;
                //Рубильники TwinBlock
                case "SelectionTwinBlock_Button":
                    await SelectionTwinBlock.getInstance(dataInXml, accessData, formSettings);
                    break;
                //Расчет обогрева
                case "TermoCalculation_Button":
                    await TermoCalculation.getInstance(formSettings, accessData);
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
                        var about = new AboutBox1();
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

        public string GetLabelText(Office.IRibbonControl control)
        {
            if (!locationDataBase)         
                return "локальная";           
            return "глобальная";
        }

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
