using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMacroAdd
{
    enum ContainerAvt
    {
        FirstLineArray,
        SecondLineArray,
        ThirdLineArray,
        FourthLineArray,
        FifthLineArray,
        SixthLineArray
    }

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
            string[] circutAvt = new string[19] { "1", "2", "3", "4", "5", "6", "8", "10", "13", "16", "20", "25", "32", "40", "50", "63", "80", "100", "125" };
            string[] kurveAvt = new string[6] { "B", "C", "D", "K", "L", "Z" };
            string[] icuAvt = new string[4] { "4,5", "6", "10", "15" };
            string[] polusAvt = new string[6] { "1", "2", "3", "4", "1N", "3N"};
            string[] vendorAvt = new string[11] { "IEK ВА47", "IEK BA47М", "IEK Armat", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DKC", "DEKraft", "Schneider", "TDM" };
            //Массивы параметров выключателей нагрузки
            string[] circutVn = new string[10] { "16", "20", "25", "32", "40", "50", "63", "80", "100", "125" };
            string[] polusVn = new string[4] { "1", "2", "3", "4" };
            string[] vendorVn = new string[8] { "IEK", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DEKraft", "Schneider", "TDM" };

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
        /// <summary>
        /// Проверка элемента в базе данных
        /// </summary>
        /// <param name="rowsCheck"></param>
        /// <param name="writeExcel"></param>
        private void CheckData(int rowsCheck)
        {
            PictureBox[] pictures = default;

            CheckBox[] checks= default;

            ComboBox[,] comboBoxes= default;

            if (tabControl1.SelectedTab == tabPage1)
            {
                 pictures = PictureBoxesCircutBreak();

                 checks = CheckBoxArrayCircutBreak();

                 comboBoxes = ComboBoxArrayCircutBreaker();
            }
            else if (tabControl1.SelectedTab == tabPage2)
            {
                pictures = PictureBoxesSwitch();

                checks = CheckBoxArraySwitch();

                comboBoxes = ComboBoxArraySwitch();
            }

            // Если стоит галочка в CheckBox, то условие истина
            if (checks[rowsCheck].Checked)
            {
                string setRequest = default;
                if (tabControl1.SelectedTab == tabPage1)
                {
                    // создаем кортеж для сборки запроса SQL
                    var tuple = (cirkut: comboBoxes[rowsCheck, 0].SelectedItem.ToString(),
                                  kurve: comboBoxes[rowsCheck, 1].SelectedItem.ToString(),
                                    icu: comboBoxes[rowsCheck, 2].SelectedItem.ToString(),
                                  polus: comboBoxes[rowsCheck, 3].SelectedItem.ToString(),
                                 vendor: comboBoxes[rowsCheck, 4].SelectedItem.ToString());
                    // Переменная запроса SQL
                    setRequest = String.Format("SELECT {0} FROM modul WHERE s_in = '{1}' AND kurve = '{2}' AND icu = '{3}' AND quantity = '{4}';",
                                                   Replace.FuncReplece(tuple.vendor), tuple.cirkut, tuple.kurve, tuple.icu, tuple.polus);
                }
                else if (tabControl1.SelectedTab == tabPage2)
                {
                    var tuple = (cirkut: comboBoxes[rowsCheck, 0].SelectedItem.ToString(),                             
                                  polus: comboBoxes[rowsCheck, 1].SelectedItem.ToString(),
                                 vendor: comboBoxes[rowsCheck, 2].SelectedItem.ToString());
                    // Переменная запроса SQL
                    setRequest = String.Format("SELECT {0} FROM switch WHERE s_in = '{1}' AND quantity = '{2}';",
                                                   Replace.FuncReplece(tuple.vendor), tuple.cirkut, tuple.polus);
                }

                //Обращение к БД в новом потоке, что бы не тормозил интерфейс
                new Thread(() =>
                {
                    //Работа с базой данных
                    DBConect classDB = new DBConect();
                    classDB.OpenDB();

                    string getArticle = classDB.RequestDB(setRequest, 0) ?? "@";

                    if (getArticle != "@")
                    {
                        //Запуск через делегат, т.к. другой поток
                        this.Invoke((MethodInvoker)delegate ()
                        {
                            pictures[rowsCheck].BackColor = Color.Green;
                        });
                    }
                    else
                    {
                        //Запуск через делегат, т.к. другой поток
                        this.Invoke((MethodInvoker)delegate ()
                        {
                            pictures[rowsCheck].BackColor = Color.IndianRed;
                        });
                    }
                    classDB.CloseDB();
                }).Start();
            }
        }
        /// <summary>
        /// Данный метод предназначен для извленчения уже заполненых данных из БД и заппуска метода заполнения листа Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            PictureBox[] pictures = default;

            CheckBox[] checks = default;

            ComboBox[,] comboBoxes = default;

            TextBox[] texts  = default;

            if (tabControl1.SelectedTab == tabPage1)
            {
                pictures = PictureBoxesCircutBreak();

                checks = CheckBoxArrayCircutBreak();

                comboBoxes = ComboBoxArrayCircutBreaker();

                texts = TextBoxesArrayCircutBreak();
            }
            else if (tabControl1.SelectedTab == tabPage2)
            {
                pictures = PictureBoxesSwitch();

                checks = CheckBoxArraySwitch();

                comboBoxes = ComboBoxArraySwitch();

                texts = TextBoxesArraySwitch();
            }

            for (int rows = 0; rows < 6; rows++)
            {
                // Если стоит галочка в CheckBox, то условие истина
                if (checks[rows].Checked)
                {
                    string setRequest = default;
                    string vendor = default;
                    
                    if (tabControl1.SelectedTab == tabPage1)
                    {
                        // создаем кортеж для сборки запроса SQL
                        var tuple = (cirkut: comboBoxes[rows, 0].SelectedItem.ToString(),
                                      kurve: comboBoxes[rows, 1].SelectedItem.ToString(),
                                        icu: comboBoxes[rows, 2].SelectedItem.ToString(),
                                      polus: comboBoxes[rows, 3].SelectedItem.ToString(),
                                     vendor: comboBoxes[rows, 4].SelectedItem.ToString());
                        // Переменная запроса SQL
                        setRequest = String.Format("SELECT {0} FROM modul WHERE s_in = '{1}' AND kurve = '{2}' AND icu = '{3}' AND quantity = '{4}';",
                                                       Replace.FuncReplece(tuple.vendor), tuple.cirkut, tuple.kurve, tuple.icu, tuple.polus);
                        vendor = tuple.vendor;
                    }
                    else if (tabControl1.SelectedTab == tabPage2)
                    {
                        var tuple = (cirkut: comboBoxes[rows, 0].SelectedItem.ToString(),
                                      polus: comboBoxes[rows, 1].SelectedItem.ToString(),
                                     vendor: comboBoxes[rows, 2].SelectedItem.ToString());
                        // Переменная запроса SQL
                        setRequest = String.Format("SELECT {0} FROM switch WHERE s_in = '{1}' AND quantity = '{2}';",
                                                       Replace.FuncReplece(tuple.vendor), tuple.cirkut, tuple.polus);
                        vendor = tuple.vendor;
                    }

                    //Работа с базой данных
                    DBConect classDB = new DBConect();
                    classDB.OpenDB();

                    string getArticle = classDB.RequestDB(setRequest, 0) ?? "@";

                    if (getArticle != "@")
                    {
                        int.TryParse(texts[rows].Text, out int quantity);
                        new WriteExcel(new DataInXml() { Vendor = vendor }, rows, getArticle, quantity, checkBox14.Checked);
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

        private PictureBox[] PictureBoxesCircutBreak()
        {
            PictureBox[] pictures = new PictureBox[6] { pictureBox1, pictureBox2, pictureBox3, pictureBox4, pictureBox5, pictureBox6 };
            return pictures;
        }

        private PictureBox[] PictureBoxesSwitch()
        {
            PictureBox[] pictures = new PictureBox[6] { pictureBox7, pictureBox8, pictureBox9, pictureBox10, pictureBox11, pictureBox12 };
            return pictures;
        }

        private TextBox[] TextBoxesArrayCircutBreak()
        {
            TextBox[] texts = new TextBox[6] { textBox1, textBox2, textBox3, textBox4, textBox5, textBox6 };
            return texts;
        }

        private TextBox[] TextBoxesArraySwitch()
        {
            TextBox[] texts = new TextBox[6] { textBox7, textBox8, textBox9, textBox10, textBox11, textBox12 };
            return texts;
        }

        private CheckBox[] CheckBoxArrayCircutBreak()
        {
            CheckBox[] checks = new CheckBox[6] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 };
            return checks;
        }

        private CheckBox[] CheckBoxArraySwitch()
        {
            CheckBox[] checks = new CheckBox[6] { checkBox7, checkBox8, checkBox9, checkBox10, checkBox11, checkBox12 };
            return checks;
        }

        private ComboBox[,] ComboBoxArrayCircutBreaker()
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

        private ComboBox[,] ComboBoxArraySwitch()
        {
            ComboBox[,] comboBoxes = new ComboBox[,]
            {
                {
                    comboBox35, comboBox32, comboBox31
                },
                {
                    comboBox40, comboBox37, comboBox36
                },
                {
                    comboBox45, comboBox42, comboBox41
                },
                {
                    comboBox50, comboBox47, comboBox46
                },
                {
                    comboBox55, comboBox52, comboBox51
                },
                {
                    comboBox60, comboBox57, comboBox56
                }
            };
            return comboBoxes;
        }



        #region line1_CircutBreak
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FirstLineArray);
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FirstLineArray);
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FirstLineArray);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FirstLineArray);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FirstLineArray);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FirstLineArray);
        }

        #endregion

        #region line2_CircutBreak

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SecondLineArray);
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SecondLineArray);
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SecondLineArray);
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SecondLineArray);
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SecondLineArray);
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SecondLineArray);
        }

        #endregion

        #region line3_CircutBreak
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.ThirdLineArray);
        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.ThirdLineArray);
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.ThirdLineArray);
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.ThirdLineArray);
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.ThirdLineArray);
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.ThirdLineArray);
        }

        #endregion

        #region line4_CircutBreak

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FourthLineArray);
        }

        private void comboBox20_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FourthLineArray);
        }

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FourthLineArray);
        }

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FourthLineArray);
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FourthLineArray);
        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FourthLineArray);
        }

        #endregion

        #region line5_CircutBreak

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FifthLineArray);
        }

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FifthLineArray);
        }

        private void comboBox24_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FifthLineArray);
        }

        private void comboBox23_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FifthLineArray);
        }

        private void comboBox22_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FifthLineArray);
        }

        private void comboBox21_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FifthLineArray);
        }


        #endregion

        #region line6_CircutBreak

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SixthLineArray);
        }

        private void comboBox30_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SixthLineArray);
        }

        private void comboBox29_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SixthLineArray);
        }

        private void comboBox28_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SixthLineArray);
        }

        private void comboBox27_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SixthLineArray);
        }

        private void comboBox26_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SixthLineArray);
        }

        #endregion

        #region line1_Switch
        private void comboBox35_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FirstLineArray);
        }

        private void comboBox32_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FirstLineArray);
        }

        private void comboBox31_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FirstLineArray);
        }

        #endregion

        #region line2_Switch
        private void comboBox40_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SecondLineArray);
        }

        private void comboBox37_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SecondLineArray);
        }

        private void comboBox36_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SecondLineArray);
        }

        #endregion

        #region line3_Switch
        private void comboBox45_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.ThirdLineArray);
        }

        private void comboBox42_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.ThirdLineArray);
        }

        private void comboBox41_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.ThirdLineArray);
        }

        #endregion

        #region line4_Switch
        private void comboBox50_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FourthLineArray);
        }

        private void comboBox47_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FourthLineArray);
        }

        private void comboBox46_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FourthLineArray);
        }

        #endregion

        #region line5_Switch
        private void comboBox55_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FifthLineArray);
        }

        private void comboBox52_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FifthLineArray);
        }

        private void comboBox51_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.FifthLineArray);
        }

        #endregion

        #region line6_Switch
        private void comboBox60_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SixthLineArray);
        }

        private void comboBox57_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SixthLineArray);
        }

        private void comboBox56_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckData((int)ContainerAvt.SixthLineArray);
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
                            
        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
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
