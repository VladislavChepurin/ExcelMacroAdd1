using ExcelMacroAdd.AccessLayer;
using ExcelMacroAdd.Forms;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.ProxyObjects;
using ExcelMacroAdd.Serializable;
using ExcelMacroAdd.Services;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;
using Path = System.IO.Path;

namespace ExcelMacroAdd
{
    public partial class MainRibbon
    {
        private readonly string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config/appSettings.json");         

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            AppSettingsDeserialize app= new AppSettingsDeserialize(jsonFilePath);
            var settings = app.GetSettingsModels();
            var resources = settings.Resources;
            var correctFontResources = settings.CorrectFontResources;
            var formSettings = settings.FormSettings;

            //Если недоступна база данных прописанная в AppSettings.json, то используется локальная
            string path;
            if (File.Exists(settings.GlobalDateBaseLocation + "BdMacro.sqlite"))
            {
                path = settings.GlobalDateBaseLocation;
                label5.Label = "глобальная";
            }
            else
            {
                path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DataLayer/DataBase");
                label5.Label = "локальная";
            }
            //Создание внедряемых зависимостей
            var dataInXml = new DataInXmlProxy(new Lazy<DataInXml>());
            var context = new AppContext(path);
            var accessData = new AccessData(context);
            //Ранняя инициализация Entity Framework
            //Будет утекать 50МБ памяти
            new Thread(() =>
            {
                if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DataLayer/DataBase/BdMacro.sqlite")))
                {
                    context.Switchs.Select(x => x.Id).ToList();
                }
            }).Start();

            // Заполнение паспортов
            button1.Click += (s, a) =>
            {
                var fillingOutThePassport = new FillingOutThePassport(resources);
                fillingOutThePassport.Start();
            };

            //Удаление формул выделеной области
            button2.Click += (s, a) => {
                var deleteFormula = new DeleteFormula();
                deleteFormula.Start();
            };

            // Удаление формул на всех листах кроме первого
            button3.Click += (s, a) =>
            {
                var deleteAllFormula = new DeleteAllFormula();
                deleteAllFormula.Start();
            };     
       
            //Корпуса щитов
            button4.Click += (s, a) => {
                var boxShield = new BoxShield(accessData, resources);
                boxShield.Start();
            };
          
            // Занесение в базу данных корпуса
            button5.Click += (s, a) => {
                var addBoxDb = new AddBoxDb(accessData, resources);
                addBoxDb.Start();
            };
            // Корректировка записей в БД
            button6.Click += (s, a) =>
            {
                var correctDb = new CorrectDb(accessData, resources);
                correctDb.Start();
            };
            //Разметка расчетов
            button7.Click += (s, a) => {
                var linker = new Linker(correctFontResources);
                linker.Start();
            };
            // Правка расчетов
            button8.Click += (s, a) =>
            {
                var editCalculation = new EditCalculation(correctFontResources);
                editCalculation.Start();
            };
            // Разметка таблицы расчетов
            button9.Click += (s, a) =>
            {
                var calculationMarkup = new CalculationMarkup(correctFontResources);
                calculationMarkup.Start();
            };
            // Разметка границ
            button10.Click += (s, a) =>
            {
                var bordersTable = new BordersTable();
                bordersTable.Start();  
            };
            // Исправление шрифтов
            button11.Click += (s, a) =>
            {
                var correctFont = new CorrectFont(correctFontResources);
                correctFont.Start();
            };

            button12.Click += async (s, a) =>
            {
                await SelectionTransformer.getInstance(dataInXml, accessData, formSettings);
            };

            button13.Click += async (s, a) =>
            {
                await SelectionTwinBlock.getInstance(dataInXml, accessData, formSettings);
            };

            button14.Click += async (s, a) =>
            {
                await TermoCalculation.getInstance(formSettings);
            };

            // Вставка формул IEK
            button20.Click += (s, a) => {
                var writeExcel = new WriteExcel(dataInXml, "Iek");
                writeExcel.Start();       
            };
            // Вставка формул EKF
            button21.Click += (s, a) => {
                var writeExcel = new WriteExcel(dataInXml, "Ekf");
                writeExcel.Start();
            };
            // Вставка формул DKC
            button22.Click += (s, a) => {
                var writeExcel = new WriteExcel(dataInXml, "Dkc");
                writeExcel.Start();
            };
            // Вставка формул KEAZ
            button23.Click += (s, a) => {
                var writeExcel = new WriteExcel(dataInXml, "Keaz");
                writeExcel.Start();   
            };
            // Вставка формул DEKraft
            button24.Click += (s, a) => {
                var writeExcel = new WriteExcel(dataInXml, "Dekraft");
                writeExcel.Start();
            };
            // Вставка формул Chint
            button25.Click += (s, a) => {
                var writeExcel = new WriteExcel(dataInXml, "Chint");
                writeExcel.Start();
            };
            // Модульные аппрараты
            button30.Click += async (s, a) =>
            {
                await SelectionCircuitBreaker.getInstance(dataInXml, accessData, formSettings);
            };
            button31.Click += async (s, a) =>
            {
                await Task.Run(() =>
                {
                    Settings fs = new Settings(dataInXml);
                    fs.ShowDialog();
                    Thread.Sleep(5000);
                });
            };

            // Окно "О программе"
            button32.Click += async (s, a) =>
            {
                await Task.Run(() =>
                {
                    var about = new AboutBox1();
                    about.ShowDialog();
                    Thread.Sleep(5000);
                });
            };

            button33.Click += (s, a) =>
            {
                System.Diagnostics.Process.Start("explorer.exe", AppDomain.CurrentDomain.BaseDirectory);
            };
            
            var getRate = new GetCurrencyTsb
            {
                CurrencyHandler = ShowCurrencyPrice
            };
            //В новом потоке запускаем метод получения данных от Центробанка
            new Thread(() =>
            {
                getRate.Start();
            }).Start();
        }

        private void ShowCurrencyPrice(double usdCurrency, double euroCurrency, double cnhCurrency)
        {
            label1.Label = "Доллар = " + usdCurrency;
            label2.Label = "ЕВРО     = " + euroCurrency;
            label3.Label = "Юань    = " + cnhCurrency;
        }
    }       
}
