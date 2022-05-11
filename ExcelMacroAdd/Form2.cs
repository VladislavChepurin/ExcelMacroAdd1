﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMacroAdd
{
    public partial class Form2 : Form
    {
        readonly int CircutIndAvt = 5; // Начальный ток автоматических выключателей
        readonly int KurveIndAvt = 1;  // Начальная кривая автоматических выключателей
        readonly int IcuIndAvt = 0;    // Начальная отключающая способность автоматических выключателей
        readonly int PolusIndAvt = 0;  // Начальная кол-во полюсов автоматических выключателей
        readonly int VendorIndAvt = 0; // Начальный вендор автолмтических выключателей

        readonly int CirkutIndVn = 0;  // Начальный ток выключателей нагрузки
        readonly int PolusIndVn = 0;  // Начальная кол-во полюсов выключателей нагрузки
        readonly int VendorIndVn = 0; // Начальный вендор выключателей нагрузки

        public Form2()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            //Массивы параметров модульных автоматов
            string[] circutAvt = new string[16] { "1", "2", "3", "4", "5", "6", "8", "10", "13", "16", "20", "25", "32", "40", "50", "63" };
            string[] kurveAvt = new string[3] { "B", "C", "D" };
            string[] icuAvt = new string[3] { "4,5", "6", "10" };
            string[] polusAvt = new string[4] { "1", "2", "3", "4" };
            string[] vendorAvt = new string[10] { "IEK ВА47", "IEK BA47М", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DKC", "DEKraft", "Schneider", "TDM" };
            //Массивы параметров выключателей нагрузки
            string[] circutVn = new string[10] { "16", "20", "25", "32", "40", "50", "63", "80", "100", "125" };
            string[] polusVn = new string[4] { "1", "2", "3", "4" };
            string[] vendorVn = new string[8] { "IEK", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DEKraft", "Schneider Electric", "TDM" };

            //Создание массивов ComboBox для автоматических выключателей
            ComboBox[] comboBoxItCircut = new ComboBox[6] { comboBox5, comboBox10, comboBox15, comboBox20, comboBox25, comboBox30 };
            ComboBox[] comboBoxItKurve = new ComboBox[6] { comboBox4, comboBox9, comboBox14, comboBox19, comboBox24, comboBox29 };
            ComboBox[] comboBoxItIcu = new ComboBox[6] { comboBox3, comboBox8, comboBox13, comboBox18, comboBox23, comboBox28 };
            ComboBox[] comboBoxItPolus = new ComboBox[6] { comboBox2, comboBox7, comboBox12, comboBox17, comboBox22, comboBox27 };
            ComboBox[] comboBoxItVendor = new ComboBox[6] { comboBox1, comboBox6, comboBox11, comboBox16, comboBox21, comboBox26 };
            //Создание массивов ComboBox для выключателей нагрузки
            ComboBox[] comboBoxItCircutVn = new ComboBox[6] { comboBox35, comboBox40, comboBox45, comboBox50, comboBox55, comboBox60 };
            ComboBox[] comboBoxItPolusVn = new ComboBox[6] { comboBox32, comboBox37, comboBox42, comboBox47, comboBox52, comboBox57 };
            ComboBox[] comboBoxItVendorVn = new ComboBox[6] { comboBox31, comboBox36, comboBox41, comboBox46, comboBox51, comboBox56 };

            for (int i = 0; i < 6; i++)
            {
                //Добавление в модульные автоматы данных тока
                comboBoxItCircut[i].Items.AddRange(circutAvt);
                comboBoxItCircut[i].SelectedIndex = CircutIndAvt;
                //Добавление в модульные автоматы данных по кривой
                comboBoxItKurve[i].Items.AddRange(kurveAvt);
                comboBoxItKurve[i].SelectedIndex = KurveIndAvt;
                //Добавление в модульные автоматы данных по макс току
                comboBoxItIcu[i].Items.AddRange(icuAvt);
                comboBoxItIcu[i].SelectedIndex = IcuIndAvt;
                //Добавление в модульные автоматы данных по полюсам
                comboBoxItPolus[i].Items.AddRange(polusAvt);
                comboBoxItPolus[i].SelectedIndex = PolusIndAvt;
                //Добавление в модульные автоматы данных по вендорам
                comboBoxItVendor[i].Items.AddRange(vendorAvt);
                comboBoxItVendor[i].SelectedIndex = VendorIndAvt;
                //Добавление в выключатели нагрузки данных тока
                comboBoxItCircutVn[i].Items.AddRange(circutVn);
                comboBoxItCircutVn[i].SelectedIndex = CirkutIndVn;
                //Добавление в выключатели нагрузки данных по полюсам
                comboBoxItPolusVn[i].Items.AddRange(polusVn);
                comboBoxItPolusVn[i].SelectedIndex = PolusIndVn;
                //Добавление в выключатели нагрузки данных по вендорам
                comboBoxItVendorVn[i].Items.AddRange(vendorVn);
                comboBoxItVendorVn[i].SelectedIndex = VendorIndVn;
            }
        }

        private void CheckData(int rowsCheck, Boolean writeExcel = false)
        {
            PictureBox[] pictures = PictureBoxes();

            CheckBox[] checks = ReturnCheckBoxArray();

            ComboBox[,] comboBoxes = ReturnComboBoxArray();
            // Если стоит галочка в CheckBox, то условие истина
            if (checks[rowsCheck].Checked)
            {
                // создаем кортеж для сборки запроса SQL
                var tuple = (cirkut: comboBoxes[rowsCheck, 0].SelectedItem.ToString(),
                              kurve: comboBoxes[rowsCheck, 1].SelectedItem.ToString(),
                                icu: comboBoxes[rowsCheck, 2].SelectedItem.ToString(),
                              polus: comboBoxes[rowsCheck, 3].SelectedItem.ToString(),
                             vendor: comboBoxes[rowsCheck, 4].SelectedItem.ToString());
                // Переменная запроса SQL
                string setRequest = String.Format("SELECT {0} FROM modul WHERE s_in = '{1}' AND kurve = '{2}' AND icu = '{3}' AND quantity = '{4}';",
                                               FuncReplece(tuple.vendor), tuple.cirkut, tuple.kurve, tuple.icu, tuple.polus);
                //Работа с базой данных
                DBConect classDB = new DBConect();
                classDB.OpenDB();

                string getArticle = classDB.RequestDB(setRequest, 0);

                if (getArticle != "@")
                {
                    pictures[rowsCheck].BackColor = Color.Green;
                }
                else
                {
                    pictures[rowsCheck].BackColor = Color.IndianRed;
                }
                classDB.CloseDB();
            }
        }

        private PictureBox[] PictureBoxes()
        {
            PictureBox[] pictures = new PictureBox[6] { pictureBox1, pictureBox2, pictureBox3, pictureBox4, pictureBox5, pictureBox6 };
            return pictures;
        }

        private TextBox[] TextBoxesArray()
        {
            TextBox[] texts = new TextBox[6] { textBox1, textBox2, textBox3, textBox4, textBox5, textBox6 };
            return texts;
        }

        private CheckBox[] ReturnCheckBoxArray()
        {
            CheckBox[] checks = new CheckBox[6] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 };
            return checks;
        }

        private ComboBox[,] ReturnComboBoxArray()
        {
            ComboBox[,] comboBoxes = new ComboBox[,]
            {
                {
                    comboBox5, comboBox4, comboBox3, comboBox2, comboBox1
                },
                {
                    comboBox10, comboBox9, comboBox8, comboBox7, comboBox6
                },
                {
                    comboBox15, comboBox14, comboBox13, comboBox12, comboBox11
                },
                {
                    comboBox20, comboBox19, comboBox18, comboBox17, comboBox16
                },
                {
                    comboBox25, comboBox24, comboBox23, comboBox22, comboBox21
                },
                {
                    comboBox30, comboBox29, comboBox28, comboBox27, comboBox26
                }
            };
            return comboBoxes;
        }

        private string FuncReplece(string mReplase)                          // Функция замены // индус заплачит от умиления IEK ВА47 - кирилица, IEK BA47М - латиница Переписать!!!
        {
            return mReplase.Replace("IEK ВА47", "iek_va47").Replace("IEK BA47М", "iek_va47m").Replace("EKF PROxima", "ekf_proxima").Replace("ABB", "abb").Replace("EKF AVERS", "ekf_avers").
                Replace("KEAZ", "keaz").Replace("DKC", "dkc").Replace("DEKraft", "dekraft").Replace("Schneider", "schneider").Replace("TDM", "tdm");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            for (int rows = 0; rows < 6; rows++)
            {
                CheckBox[] checks = ReturnCheckBoxArray();

                ComboBox[,] comboBoxes = ReturnComboBoxArray();

                TextBox[] texts = TextBoxesArray();

                // Если стоит галочка в CheckBox, то условие истина
                if (checks[rows].Checked)
                {
                    // создаем кортеж для сборки запроса SQL
                    var tuple = (cirkut: comboBoxes[rows, 0].SelectedItem.ToString(),
                                  kurve: comboBoxes[rows, 1].SelectedItem.ToString(),
                                    icu: comboBoxes[rows, 2].SelectedItem.ToString(),
                                  polus: comboBoxes[rows, 3].SelectedItem.ToString(),
                                 vendor: comboBoxes[rows, 4].SelectedItem.ToString());
                    // Переменная запроса SQL
                    string setRequest = String.Format("SELECT {0} FROM modul WHERE s_in = '{1}' AND kurve = '{2}' AND icu = '{3}' AND quantity = '{4}';",
                                                   FuncReplece(tuple.vendor), tuple.cirkut, tuple.kurve, tuple.icu, tuple.polus);
                    //Работа с базой данных
                    DBConect classDB = new DBConect();
                    classDB.OpenDB();

                    string getArticle = classDB.RequestDB(setRequest, 0);

                    if (getArticle != "@")
                    {
                        int.TryParse(texts[rows].Text, out int result);
                        _ = new WriteExcel(getArticle, tuple.vendor, rows, result, checkBox14.Checked);
                    }
                                  
                    classDB.CloseDB();

                }
            }
        }
        private async void button2_Click(object sender, EventArgs e)
        {
            await Task.Run(() =>
            {
                Form3 settings  = new Form3();
                settings.ShowDialog();
                Thread.Sleep(5000);
            });
        }

        #region line1
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
           CheckData(0);
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(0);
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(0);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(0);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(0);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(0);
        }

        #endregion

        #region line2

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            CheckData(1);
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(1);
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(1);
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(1);
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(1);
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(1);
        }

        #endregion

        #region line3
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            CheckData(2);
        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(2);
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(2);
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(2);
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(2);
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(2);
        }

        #endregion

        #region line4

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            CheckData(3);
        }

        private void comboBox20_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(3);
        }

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(3);
        }

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(3);
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(3);
        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(3);
        }

        #endregion

        #region line5

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            CheckData(4);
        }

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(4);
        }

        private void comboBox24_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(4);
        }

        private void comboBox23_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(4);
        }

        private void comboBox22_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(4);
        }

        private void comboBox21_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(4);
        }


        #endregion

        #region line6

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            CheckData(5);
        }

        private void comboBox30_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(5);
        }

        private void comboBox29_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(5);
        }

        private void comboBox28_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(5);
        }

        private void comboBox27_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(5);
        }

        private void comboBox26_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData(5);
        }


        #endregion

        #region KeyPress

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }


        #endregion

    }

}
