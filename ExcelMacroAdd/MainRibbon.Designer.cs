namespace ExcelMacroAdd
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Группа1 = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.label4 = this.Factory.CreateRibbonLabel();
            this.label5 = this.Factory.CreateRibbonLabel();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.button20 = this.Factory.CreateRibbonButton();
            this.button21 = this.Factory.CreateRibbonButton();
            this.button22 = this.Factory.CreateRibbonButton();
            this.button23 = this.Factory.CreateRibbonButton();
            this.button24 = this.Factory.CreateRibbonButton();
            this.button25 = this.Factory.CreateRibbonButton();
            this.button30 = this.Factory.CreateRibbonButton();
            this.button12 = this.Factory.CreateRibbonButton();
            this.button13 = this.Factory.CreateRibbonButton();
            this.button14 = this.Factory.CreateRibbonButton();
            this.button31 = this.Factory.CreateRibbonButton();
            this.button32 = this.Factory.CreateRibbonButton();
            this.button33 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Группа1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group1.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.Группа1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "Абиэлт";
            this.tab1.Name = "tab1";
            // 
            // Группа1
            // 
            this.Группа1.Items.Add(this.button1);
            this.Группа1.Items.Add(this.button2);
            this.Группа1.Items.Add(this.button3);
            this.Группа1.Items.Add(this.button4);
            this.Группа1.Items.Add(this.button5);
            this.Группа1.Items.Add(this.button6);
            this.Группа1.Items.Add(this.button7);
            this.Группа1.Items.Add(this.button8);
            this.Группа1.Items.Add(this.separator1);
            this.Группа1.Items.Add(this.button9);
            this.Группа1.Items.Add(this.button10);
            this.Группа1.Items.Add(this.button11);
            this.Группа1.Label = "Базовые макросы";
            this.Группа1.Name = "Группа1";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // group3
            // 
            this.group3.Items.Add(this.button20);
            this.group3.Items.Add(this.button21);
            this.group3.Items.Add(this.button22);
            this.group3.Items.Add(this.button23);
            this.group3.Items.Add(this.button24);
            this.group3.Items.Add(this.button25);
            this.group3.Items.Add(this.separator2);
            this.group3.Items.Add(this.button30);
            this.group3.Items.Add(this.button12);
            this.group3.Items.Add(this.button13);
            this.group3.Items.Add(this.button14);
            this.group3.Items.Add(this.button31);
            this.group3.Label = "Макросы для расчетов";
            this.group3.Name = "group3";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button32);
            this.group1.Items.Add(this.button33);
            this.group1.Name = "group1";
            // 
            // group4
            // 
            this.group4.Items.Add(this.label4);
            this.group4.Items.Add(this.label5);
            this.group4.Name = "group4";
            // 
            // label4
            // 
            this.label4.Label = "База данных";
            this.label4.Name = "label4";
            // 
            // label5
            // 
            this.label5.Label = "Не готова";
            this.label5.Name = "label5";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Заполнить паспорта";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "Удалить формулы";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "Удалить все формулы";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.Label = "Корпуса щитов";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
            this.button5.Label = "Корпуса в базу";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Image = ((System.Drawing.Image)(resources.GetObject("button6.Image")));
            this.button6.Label = "Исправить запись БД";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            // 
            // button7
            // 
            this.button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button7.Image = global::ExcelMacroAdd.Properties.Resources._302;
            this.button7.Label = "Разметка листов";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            // 
            // button8
            // 
            this.button8.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button8.Image = global::ExcelMacroAdd.Properties.Resources._305;
            this.button8.Label = "Причесать расчет";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            // 
            // button9
            // 
            this.button9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button9.Image = global::ExcelMacroAdd.Properties.Resources._970;
            this.button9.Label = "Расчет";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            // 
            // button10
            // 
            this.button10.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button10.Image = ((System.Drawing.Image)(resources.GetObject("button10.Image")));
            this.button10.Label = "Границы";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            // 
            // button11
            // 
            this.button11.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button11.Image = global::ExcelMacroAdd.Properties.Resources._902;
            this.button11.Label = "Шрифт";
            this.button11.Name = "button11";
            this.button11.ShowImage = true;
            // 
            // button20
            // 
            this.button20.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button20.Image = global::ExcelMacroAdd.Properties.Resources.iek;
            this.button20.Label = "Формула ВПР IEK";
            this.button20.Name = "button20";
            this.button20.ShowImage = true;
            // 
            // button21
            // 
            this.button21.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button21.Image = global::ExcelMacroAdd.Properties.Resources.ekf;
            this.button21.Label = "Формула ВПР EKF";
            this.button21.Name = "button21";
            this.button21.ShowImage = true;
            // 
            // button22
            // 
            this.button22.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button22.Image = global::ExcelMacroAdd.Properties.Resources.dkc;
            this.button22.Label = "Формула ВПР DKC";
            this.button22.Name = "button22";
            this.button22.ShowImage = true;
            // 
            // button23
            // 
            this.button23.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button23.Image = global::ExcelMacroAdd.Properties.Resources.keaz;
            this.button23.Label = "Формула ВПР KEAZ";
            this.button23.Name = "button23";
            this.button23.ShowImage = true;
            // 
            // button24
            // 
            this.button24.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button24.Image = global::ExcelMacroAdd.Properties.Resources._560;
            this.button24.Label = "Формула ВПР DEK";
            this.button24.Name = "button24";
            this.button24.ShowImage = true;
            // 
            // button25
            // 
            this.button25.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button25.Image = global::ExcelMacroAdd.Properties.Resources.chint;
            this.button25.Label = "Формула ВПР CHINT";
            this.button25.Name = "button25";
            this.button25.ShowImage = true;
            // 
            // button30
            // 
            this.button30.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button30.Image = global::ExcelMacroAdd.Properties.Resources._501;
            this.button30.Label = "Модульные аппараты";
            this.button30.Name = "button30";
            this.button30.ShowImage = true;
            // 
            // button12
            // 
            this.button12.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button12.Image = global::ExcelMacroAdd.Properties.Resources._503;
            this.button12.Label = "Трансформатор тока";
            this.button12.Name = "button12";
            this.button12.ShowImage = true;
            // 
            // button13
            // 
            this.button13.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button13.Image = ((System.Drawing.Image)(resources.GetObject("button13.Image")));
            this.button13.Label = "Рубильники TwinBlock";
            this.button13.Name = "button13";
            this.button13.ShowImage = true;
            // 
            // button14
            // 
            this.button14.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button14.Image = global::ExcelMacroAdd.Properties.Resources._505;
            this.button14.Label = "Расчет обогрева";
            this.button14.Name = "button14";
            this.button14.ShowImage = true;
            // 
            // button31
            // 
            this.button31.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button31.Image = ((System.Drawing.Image)(resources.GetObject("button31.Image")));
            this.button31.Label = "Настройка формул";
            this.button31.Name = "button31";
            this.button31.ShowImage = true;
            // 
            // button32
            // 
            this.button32.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button32.Image = global::ExcelMacroAdd.Properties.Resources._201;
            this.button32.Label = "About";
            this.button32.Name = "button32";
            this.button32.ShowImage = true;
            // 
            // button33
            // 
            this.button33.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button33.Image = global::ExcelMacroAdd.Properties.Resources.Open;
            this.button33.Label = "Открыть папку";
            this.button33.Name = "button33";
            this.button33.ShowImage = true;
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Группа1.ResumeLayout(false);
            this.Группа1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Группа1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button32;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button30;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button31;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button20;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button21;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button22;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button23;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button24;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button33;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label4;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button25;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button14;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon Ribbon1
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
