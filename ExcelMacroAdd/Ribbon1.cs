using ExcelMacroAdd.Forms;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Services;
using ExcelMacroAdd.Servises;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelMacroAdd
{
    public partial class Ribbon1
    {     
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {             
            Lazy<DataInXml> dataInXml = new Lazy<DataInXml>();
            DBConectProxy dBConect = new DBConectProxy(new Lazy<DBConect>());

            // Заполнение паспортов
            button1.Click += (s, a) =>
            {
                FillingOutThePassport fillingOutThePassport = new FillingOutThePassport(dBConect);
                fillingOutThePassport.Start();
            };

            //Удаление формул выделеной области
            button2.Click += (s, a) => {
                DeleteFormula deleteFormula = new DeleteFormula();
                deleteFormula.Start();
            };

            // Удаление формул на всех листах кроме первого
            button3.Click += (s, a) =>
            {
                DeleteAllFormula deleteAllFormula = new DeleteAllFormula();
                deleteAllFormula.Start();
            };     
       
            //Корпуса щитов
            button4.Click += (s, a) => {
                BoxShield boxShield = new BoxShield(dBConect);
                boxShield.Start();
            };
          
            // Занесение в базу данных корпуса
            button5.Click += (s, a) => {
                AddBoxDB addBoxDB = new AddBoxDB(dBConect);
                addBoxDB.Start();
            };
            // Корректировка записей в БД
            button6.Click += (s, a) =>
            {
                CorectDB corectDB = new CorectDB(dBConect);
                corectDB.Start();
            };
            //Разметка расчетов
            button7.Click += (s, a) => {
                Linker linker = new Linker();
                linker.Start();
            };
            // Правка расчетов
            button8.Click += (s, a) =>
            {
                EditCalculation editCalculation = new EditCalculation();
                editCalculation.Start();
            };
            // Разметка таблицы расчетов
            button9.Click += (s, a) =>
            {
                CalculationMarkup calculationMarkup = new CalculationMarkup();
                calculationMarkup.Start();
            };
            // Разметка границ
            button10.Click += (s, a) =>
            {
                BordersTable bordersTable = new BordersTable();
                bordersTable.Start();  
            };
            // Исправление шрифтов
            button11.Click += (s, a) =>
            {
                CorrectFont correctFont = new CorrectFont();
                correctFont.Start();
            };               
            // Вставка формул IEK
            button20.Click += (s, a) => {
                WriteExcel writeExcel = new WriteExcel(dataInXml, "IEK");
                writeExcel.Start();       
            };
            // Вставка формул EKF
            button21.Click += (s, a) => {
                WriteExcel writeExcel = new WriteExcel(dataInXml, "EKF");
                writeExcel.Start();
            };
            // Вставка формул DKC
            button22.Click += (s, a) => {
                WriteExcel writeExcel = new WriteExcel(dataInXml, "DKC");
                writeExcel.Start();
            };
            // Вставка формул KEAZ
            button23.Click += (s, a) => {
                WriteExcel writeExcel = new WriteExcel(dataInXml, "KEAZ");
                writeExcel.Start();   
            };
            // Вставка формул DEKraft
            button24.Click += (s, a) => {
                WriteExcel writeExcel = new WriteExcel(dataInXml, "DEKraft");
                writeExcel.Start();
            };
            // Модульные аппрараты
            button30.Click += async (s, a) =>
            {
                await Task.Run(() =>
                {
                    Form2 fs = new Form2(dBConect, dataInXml);
                    fs.ShowDialog();
                    Thread.Sleep(5000);
                });
            };
            // Настройки
            button31.Click += async (s, a) =>
            {
                await Task.Run(() =>
                {
                    Form3 fs = new Form3(dataInXml);
                    fs.ShowDialog();
                    Thread.Sleep(5000);
                });
            };
            // Окно "О программе"
            button32.Click += async (s, a) =>
            {
                await Task.Run(() =>
                {
                    AboutBox1 about = new AboutBox1();
                    about.ShowDialog();
                    Thread.Sleep(5000);
                });
            };            

            GetValuteTSB getRate = new GetValuteTSB
            {
                ValuteUSDHandler = ShowValitePrice
            };
            //В новом потоке запускаем метод получения данных от Центробанка
            new Thread(() =>
            {
                getRate.Start();
                //Thread.Sleep(100);
            }).Start();          
        }
            
        private void ShowValitePrice(double usdValute, double evroValute, double cnhValute)
        {
            this.label1.Label = "Доллар = " + usdValute;
            this.label2.Label = "ЕВРО     = " + evroValute;
            this.label3.Label = "Юань    = " + cnhValute;
        }              
    }       
}
