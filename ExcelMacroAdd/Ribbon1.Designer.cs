namespace ExcelMacroAdd
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Группа1 = this.Factory.CreateRibbonGroup();
            this.button5 = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button19 = this.Factory.CreateRibbonButton();
            this.button17 = this.Factory.CreateRibbonButton();
            this.button18 = this.Factory.CreateRibbonButton();
            this.button13 = this.Factory.CreateRibbonButton();
            this.button14 = this.Factory.CreateRibbonButton();
            this.button15 = this.Factory.CreateRibbonButton();
            this.button16 = this.Factory.CreateRibbonButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.button12 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Группа1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.Группа1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "МАКРОСЫ";
            this.tab1.Name = "tab1";
            // 
            // Группа1
            // 
            this.Группа1.Items.Add(this.button6);
            this.Группа1.Items.Add(this.button1);
            this.Группа1.Items.Add(this.button8);
            this.Группа1.Items.Add(this.button3);
            this.Группа1.Items.Add(this.button4);
            this.Группа1.Items.Add(this.button9);
            this.Группа1.Items.Add(this.button2);
            this.Группа1.Items.Add(this.button7);
            this.Группа1.Items.Add(this.button5);
            this.Группа1.Items.Add(this.separator1);
            this.Группа1.Items.Add(this.button19);
            this.Группа1.Items.Add(this.button17);
            this.Группа1.Items.Add(this.button18);
            this.Группа1.Label = "Базовые макросы";
            this.Группа1.Name = "Группа1";
            // 
            // button5
            // 
            this.button5.Label = "";
            this.button5.Name = "button5";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // group3
            // 
            this.group3.Items.Add(this.button13);
            this.group3.Items.Add(this.button14);
            this.group3.Items.Add(this.button15);
            this.group3.Items.Add(this.button16);
            this.group3.Items.Add(this.separator2);
            this.group3.Items.Add(this.button11);
            this.group3.Items.Add(this.button12);
            this.group3.Label = "Макросы для расчетов";
            this.group3.Name = "group3";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.label1);
            this.group2.Items.Add(this.label2);
            this.group2.Items.Add(this.label3);
            this.group2.Label = "Курсы валют ЦБ РФ";
            this.group2.Name = "group2";
            // 
            // label1
            // 
            this.label1.Label = "label1";
            this.label1.Name = "label1";
            // 
            // label2
            // 
            this.label2.Label = "label2";
            this.label2.Name = "label2";
            // 
            // label3
            // 
            this.label3.Label = "label3";
            this.label3.Name = "label3";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button10);
            this.group1.Name = "group1";
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Image = global::ExcelMacroAdd.Properties.Resources._306;
            this.button6.Label = "Заполнить паспорта";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::ExcelMacroAdd.Properties.Resources._308;
            this.button1.Label = "Удалить формулы";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button8
            // 
            this.button8.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button8.Image = global::ExcelMacroAdd.Properties.Resources._301;
            this.button8.Label = "Удалить все формулы";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = global::ExcelMacroAdd.Properties.Resources._303;
            this.button3.Label = "Корпуса щитов";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.Label = "Корпуса в базу";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // button9
            // 
            this.button9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button9.Image = global::ExcelMacroAdd.Properties.Resources._307;
            this.button9.Label = "Исправить запись БД";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button9_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::ExcelMacroAdd.Properties.Resources._302;
            this.button2.Label = "Разметка листов";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button7
            // 
            this.button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button7.Image = global::ExcelMacroAdd.Properties.Resources._305;
            this.button7.Label = "Причесать расчет";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // button19
            // 
            this.button19.Image = global::ExcelMacroAdd.Properties.Resources._970;
            this.button19.Label = "Расчет";
            this.button19.Name = "button19";
            this.button19.ShowImage = true;
            this.button19.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button19_Click);
            // 
            // button17
            // 
            this.button17.Image = global::ExcelMacroAdd.Properties.Resources._901;
            this.button17.Label = "Границы";
            this.button17.Name = "button17";
            this.button17.ShowImage = true;
            this.button17.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button17_Click);
            // 
            // button18
            // 
            this.button18.Image = global::ExcelMacroAdd.Properties.Resources._902;
            this.button18.Label = "Шрифт";
            this.button18.Name = "button18";
            this.button18.ShowImage = true;
            this.button18.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button18_Click);
            // 
            // button13
            // 
            this.button13.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button13.Image = global::ExcelMacroAdd.Properties.Resources.iek;
            this.button13.Label = "Формула ВПР IEK";
            this.button13.Name = "button13";
            this.button13.ShowImage = true;
            this.button13.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button13_Click);
            // 
            // button14
            // 
            this.button14.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button14.Image = global::ExcelMacroAdd.Properties.Resources.ekf;
            this.button14.Label = "Формула ВПР EKF";
            this.button14.Name = "button14";
            this.button14.ShowImage = true;
            this.button14.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button14_Click);
            // 
            // button15
            // 
            this.button15.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button15.Image = global::ExcelMacroAdd.Properties.Resources.dkc;
            this.button15.Label = "Формула ВПР DKC";
            this.button15.Name = "button15";
            this.button15.ShowImage = true;
            this.button15.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button15_Click);
            // 
            // button16
            // 
            this.button16.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button16.Image = global::ExcelMacroAdd.Properties.Resources.keaz;
            this.button16.Label = "Формула ВПР KEAZ";
            this.button16.Name = "button16";
            this.button16.ShowImage = true;
            this.button16.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button16_Click);
            // 
            // button11
            // 
            this.button11.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button11.Image = global::ExcelMacroAdd.Properties.Resources._501;
            this.button11.Label = "Модульные аппараты";
            this.button11.Name = "button11";
            this.button11.ShowImage = true;
            this.button11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button11_Click);
            // 
            // button12
            // 
            this.button12.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button12.Image = global::ExcelMacroAdd.Properties.Resources._502;
            this.button12.Label = "Settings";
            this.button12.Name = "button12";
            this.button12.ShowImage = true;
            this.button12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button12_Click);
            // 
            // button10
            // 
            this.button10.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button10.Image = global::ExcelMacroAdd.Properties.Resources._201;
            this.button10.Label = "About";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            this.button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button10_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Группа1.ResumeLayout(false);
            this.Группа1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Группа1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button14;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button15;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button16;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button17;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button18;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button19;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
