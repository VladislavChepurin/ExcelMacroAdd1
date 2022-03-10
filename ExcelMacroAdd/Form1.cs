using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMacroAdd
{
    public partial class Form1 : Form 
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Массивы параметров модульных автоматов
            string[] circut_boxAvt = new string[16] { "1", "2", "3", "4", "5", "6", "8", "10", "13", "16", "20", "25", "32", "40", "50", "63" };
            string[] kurve_boxAvt = new string[3] { "B", "C", "D" };
            string[] icu_boxAvt = new string[3] { "4,5kA", "6kA", "10kA" };
            string[] polus_boxAvt = new string[4] { "1", "2", "3", "4" };
            string[] vendor_boxAvt = new string[10] { "IEK ВА47", "IEK ВА47М", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DKC", "DECraft", "Schneider Electric", "TDM" };
            //Массивы параметров выключателей нагрузки
            string[] circut_boxVn = new string[10] { "16", "20", "25", "32", "40", "50", "63", "80", "100", "125"};
            string[] polus_boxVn = new string[4] { "1", "2", "3", "4" };
            string[] vendor_boxVn = new string[8] { "IEK", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DECraft", "Schneider Electric", "TDM" };
            //Добавление в модульные автоматы данных тока
            comboBox5.Items.AddRange(circut_boxAvt);
            comboBox10.Items.AddRange(circut_boxAvt);
            comboBox15.Items.AddRange(circut_boxAvt);
            comboBox20.Items.AddRange(circut_boxAvt);
            comboBox25.Items.AddRange(circut_boxAvt);
            comboBox30.Items.AddRange(circut_boxAvt);
            //Добавление в модульные автоматы данных по кривой
            comboBox4.Items.AddRange(kurve_boxAvt);
            comboBox9.Items.AddRange(kurve_boxAvt);
            comboBox14.Items.AddRange(kurve_boxAvt);
            comboBox19.Items.AddRange(kurve_boxAvt);
            comboBox24.Items.AddRange(kurve_boxAvt);
            comboBox29.Items.AddRange(kurve_boxAvt);
            //Добавление в модульные автоматы данных по макс току
            comboBox3.Items.AddRange(icu_boxAvt);
            comboBox8.Items.AddRange(icu_boxAvt);
            comboBox13.Items.AddRange(icu_boxAvt);
            comboBox18.Items.AddRange(icu_boxAvt);
            comboBox23.Items.AddRange(icu_boxAvt);
            comboBox28.Items.AddRange(icu_boxAvt);
            //Добавление в модульные автоматы данных по полюсам
            comboBox2.Items.AddRange(polus_boxAvt);
            comboBox7.Items.AddRange(polus_boxAvt);
            comboBox12.Items.AddRange(polus_boxAvt);
            comboBox17.Items.AddRange(polus_boxAvt);
            comboBox22.Items.AddRange(polus_boxAvt);
            comboBox27.Items.AddRange(polus_boxAvt);
            //Добавление в модульные автоматы данных по вендорам
            comboBox1.Items.AddRange(vendor_boxAvt);
            comboBox6.Items.AddRange(vendor_boxAvt);
            comboBox11.Items.AddRange(vendor_boxAvt);
            comboBox16.Items.AddRange(vendor_boxAvt);
            comboBox21.Items.AddRange(vendor_boxAvt);
            comboBox26.Items.AddRange(vendor_boxAvt);
            //Добавление в выключатели нагрузки данных тока
            comboBox35.Items.AddRange(circut_boxVn);
            comboBox40.Items.AddRange(circut_boxVn);
            comboBox45.Items.AddRange(circut_boxVn);
            comboBox50.Items.AddRange(circut_boxVn);
            comboBox55.Items.AddRange(circut_boxVn);
            comboBox60.Items.AddRange(circut_boxVn);
            //Добавление в выключатели нагрузки данных по полюсам
            comboBox32.Items.AddRange(polus_boxVn);
            comboBox37.Items.AddRange(polus_boxVn);
            comboBox42.Items.AddRange(polus_boxVn);
            comboBox47.Items.AddRange(polus_boxVn);
            comboBox52.Items.AddRange(polus_boxVn);
            comboBox57.Items.AddRange(polus_boxVn);
            //Добавление в выключатели нагрузки данных по вендорам
            comboBox31.Items.AddRange(vendor_boxVn);
            comboBox36.Items.AddRange(vendor_boxVn);
            comboBox41.Items.AddRange(vendor_boxVn);
            comboBox46.Items.AddRange(vendor_boxVn);
            comboBox51.Items.AddRange(vendor_boxVn);
            comboBox56.Items.AddRange(vendor_boxVn);
        }
    }
}
