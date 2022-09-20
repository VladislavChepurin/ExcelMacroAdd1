using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Switch = ExcelMacroAdd.DataLayer.Entity.Switch;

namespace ExcelMacroAdd.Forms
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

    internal partial class Form2 : Form
    {
        readonly int CircutIndAvt = 5; // Начальный ток автоматических выключателей
        readonly int KurveIndAvt = 1;  // Начальная кривая автоматических выключателей
        readonly int IcuIndAvt = 0;    // Начальная отключающая способность автоматических выключателей
        readonly int PolusIndAvt = 0;  // Начальная кол-во полюсов автоматических выключателей
        readonly int VendorIndAvt = 0; // Начальный вендор автолмтических выключателей

        readonly int CirkutIndVn = 0;  // Начальный ток выключателей нагрузки
        readonly int PolusIndVn = 0;  // Начальная кол-во полюсов выключателей нагрузки
        readonly int VendorIndVn = 0; // Начальный вендор выключателей нагрузки

        private readonly IDataInXml dataInXml;
        private readonly IResourcesForm2 resources;
        private readonly IForm2Data accessData;

        internal Form2(IForm2Data accessData, IDataInXml dataInXml, IResourcesForm2 resources)
        {
            this.accessData = accessData;
            this.dataInXml = dataInXml;
            this.resources = resources;
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //Массивы параметров модульных автоматов
            string[] circuitBreakerCurrent = resources.CircuitBreakerCurrent;
            string[] circuitBreakerCurve = resources.CircuitBreakerCurve;
            string[] maxCircuitBreakerCurrent = resources.MaxCircuitBreakerCurrent;
            string[] amountOfPolesCircuitBreaker = resources.AmountOfPolesCircuitBreaker;
            string[] circuitBreakerVendor = resources.CircuitBreakerVendor;
            //Массивы параметров выключателей нагрузки
            string[] loadSwitchCurrent = resources.LoadSwitchCurrent;
            string[] amountOfPolesLoadSwitch = resources.AmountOfPolesLoadSwitch;
            string[] loadSwitchVendor = resources.LoadSwitchVendor;

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
                comboBoxItCircut[i].Items.AddRange(circuitBreakerCurrent);
                comboBoxItCircut[i].SelectedIndex = CircutIndAvt;
                //Добавление в модульные автоматы данных по кривой
                comboBoxItKurve[i].Items.AddRange(circuitBreakerCurve);
                comboBoxItKurve[i].SelectedIndex = KurveIndAvt;
                //Добавление в модульные автоматы данных по макс току
                comboBoxItIcu[i].Items.AddRange(maxCircuitBreakerCurrent);
                comboBoxItIcu[i].SelectedIndex = IcuIndAvt;
                //Добавление в модульные автоматы данных по полюсам
                comboBoxItPolus[i].Items.AddRange(amountOfPolesCircuitBreaker);
                comboBoxItPolus[i].SelectedIndex = PolusIndAvt;
                //Добавление в модульные автоматы данных по вендорам
                comboBoxItVendor[i].Items.AddRange(circuitBreakerVendor);
                comboBoxItVendor[i].SelectedIndex = VendorIndAvt;
                //Добавление в выключатели нагрузки данных тока
                comboBoxItCircutVn[i].Items.AddRange(loadSwitchCurrent);
                comboBoxItCircutVn[i].SelectedIndex = CirkutIndVn;
                //Добавление в выключатели нагрузки данных по полюсам
                comboBoxItPolusVn[i].Items.AddRange(amountOfPolesLoadSwitch);
                comboBoxItPolusVn[i].SelectedIndex = PolusIndVn;
                //Добавление в выключатели нагрузки данных по вендорам
                comboBoxItVendorVn[i].Items.AddRange(loadSwitchVendor);
                comboBoxItVendorVn[i].SelectedIndex = VendorIndVn;
            }
        }

        private async void CheckDataCircutBreakAsync(int rowsCheck)
        {
            PictureBox[] pictures = PictureBoxesCircutBreak();
            CheckBox[] checks = CheckBoxArrayCircutBreak();
            ComboBox[,] comboBoxes = ComboBoxArrayCircutBreak();
            // Если стоит галочка в CheckBox, то условие истина
            if (!checks[rowsCheck].Checked)
            {
                return;
            }

            string current = comboBoxes[rowsCheck, 0].SelectedItem.ToString();
            string kurve = comboBoxes[rowsCheck, 1].SelectedItem.ToString();
            string maxCurrent = comboBoxes[rowsCheck, 2].SelectedItem.ToString();
            string polus = comboBoxes[rowsCheck, 3].SelectedItem.ToString();
            string vendor = GetDisconaryVendor()[comboBoxes[rowsCheck, 4].SelectedItem.ToString()];

            try
            {
                var  modulses = await accessData.GetEntityModul(current, kurve, maxCurrent, polus);        

                if (modulses is null)
                {
                    pictures[rowsCheck].BackColor = Color.IndianRed;
                    return;
                }

                Type myType = typeof(Modul);
                // получаем свойство
                var articleProp = myType.GetProperty(vendor);
                // получаем значение свойства
                var article = articleProp?.GetValue(modulses);

                if (String.IsNullOrEmpty(article.ToString()))
                {
                    pictures[rowsCheck].BackColor = Color.IndianRed;
                }
                else
                {
                    pictures[rowsCheck].BackColor = Color.Green;
                }
            }
            catch (DataException)
            {
                MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                    "Ошибка базы данных");
            }
            catch (Exception e)
            {
                MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {e.Message}",
                    "Ошибка базы данных");
            }
        }

        private async void CheckDataSwitchAsync(int rowsCheck)
        {
            PictureBox[] pictures = PictureBoxesSwitch();
            CheckBox[] checks = CheckBoxArraySwitch();
            ComboBox[,] comboBoxes = ComboBoxArraySwitch();

            if (!checks[rowsCheck].Checked)
            {
                return;
            }

            string current = comboBoxes[rowsCheck, 0].SelectedItem.ToString();
            string polus = comboBoxes[rowsCheck, 1].SelectedItem.ToString();
            string vendor = GetDisconaryVendor()[comboBoxes[rowsCheck, 2].SelectedItem.ToString()];

            try
            {
                var switches = await accessData.GetEntitySwitch(current, polus);             

                if (switches is null)
                {
                    pictures[rowsCheck].BackColor = Color.IndianRed;
                    return;
                }

                Type myType = typeof(Switch);
                // получаем свойство
                var articleProp = myType.GetProperty(vendor);
                // получаем значение свойства
                var article = articleProp?.GetValue(switches);

                if (String.IsNullOrEmpty(article.ToString()))
                {
                    pictures[rowsCheck].BackColor = Color.IndianRed;
                }
                else
                {
                    pictures[rowsCheck].BackColor = Color.Green;
                }
            }
            catch (DataException)
            {
                MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                    "Ошибка базы данных");
            }
            catch (Exception e)
            {
                MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {e.Message}",
                    "Ошибка базы данных");
            }
        }

        /// <summary>
        /// Данный метод предназначен для извленчения уже заполненых данных из БД и запуска метода заполнения листа Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage1)
            {
                CreateFillInCircutBreakAsync();
            }
            else if (tabControl1.SelectedTab == tabPage2)
            {
                CreateFillInSwitchAsync();
            }
        }

        public async void CreateFillInCircutBreakAsync()
        {
            CheckBox[] checks = CheckBoxArrayCircutBreak();
            ComboBox[,] comboBoxes = ComboBoxArrayCircutBreak();
            TextBox[] texts = TextBoxesArrayCircutBreak();

            for (int rows = 0; rows < 6; rows++)
            {
                // Если стоит галочка в CheckBox, то условие истина
                if (checks[rows].Checked)
                {
                    string current = comboBoxes[rows, 0].SelectedItem.ToString();
                    string kurve = comboBoxes[rows, 1].SelectedItem.ToString();
                    string maxCurrent = comboBoxes[rows, 2].SelectedItem.ToString();
                    string polus = comboBoxes[rows, 3].SelectedItem.ToString();
                    string vendor = GetDisconaryVendor()[comboBoxes[rows, 4].SelectedItem.ToString()];

                    try
                    {
                        var modulses = await accessData.GetEntityModul(current, kurve, maxCurrent, polus);

                        Type myType = typeof(Modul);
                        // получаем свойство
                        var articleProp = myType.GetProperty(vendor);
                        // получаем значение свойства
                        var article = articleProp?.GetValue(modulses);

                        if (String.IsNullOrEmpty(article.ToString()))
                        {
                            continue;
                        }

                        int.TryParse(texts[rows].Text, out int quantity);
                        WriteExcel writeExcel = new WriteExcel(dataInXml, vendor, rows, article.ToString(), quantity);
                        writeExcel.Start();
                    }
                    catch (DataException)
                    {
                        MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                            "Ошибка базы данных");
                    }
                    catch (Exception e)
                    {
                        MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {e.Message}",
                            "Ошибка базы данных");
                    }
                }
            }
        }

        public async void CreateFillInSwitchAsync()
        {
            CheckBox[] checks = CheckBoxArraySwitch();
            ComboBox[,] comboBoxes = ComboBoxArraySwitch();
            TextBox[] texts = TextBoxesArraySwitch();

            for (int rows = 0; rows < 6; rows++)
            {
                // Если стоит галочка в CheckBox, то условие истина
                if (checks[rows].Checked)
                {
                    string current = comboBoxes[rows, 0].SelectedItem.ToString();
                    string polus = comboBoxes[rows, 1].SelectedItem.ToString();
                    string vendor = GetDisconaryVendor()[comboBoxes[rows, 2].SelectedItem.ToString()];

                    try
                    {
                        var switches = await accessData.GetEntitySwitch(current, polus);

                        Type myType = typeof(Switch);
                        // получаем свойство
                        var articleProp = myType.GetProperty(vendor);
                        // получаем значение свойства
                        var article = articleProp?.GetValue(switches);

                        if (String.IsNullOrEmpty(article.ToString()))
                        {
                            continue;
                        }

                        int.TryParse(texts[rows].Text, out int quantity);
                        WriteExcel writeExcel = new WriteExcel(dataInXml, vendor, rows, article.ToString(), quantity);
                        writeExcel.Start();

                    }
                    catch (DataException)
                    {
                        MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                            "Ошибка базы данных");
                    }
                    catch (Exception e)
                    {
                        MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {e.Message}",
                            "Ошибка базы данных");
                    }
                }
            }
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            await Task.Run(() =>
            {
                Form3 settings = new Form3(dataInXml);
                settings.ShowDialog();
                Thread.Sleep(5000);
            });
        }
        protected PictureBox[] PictureBoxesCircutBreak()
        {
            PictureBox[] pictures = new PictureBox[6] { pictureBox1, pictureBox2, pictureBox3, pictureBox4, pictureBox5, pictureBox6 };
            return pictures;
        }
        protected PictureBox[] PictureBoxesSwitch()
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
        protected CheckBox[] CheckBoxArraySwitch()
        {
            CheckBox[] checks = new CheckBox[6] { checkBox7, checkBox8, checkBox9, checkBox10, checkBox11, checkBox12 };
            return checks;
        }
        private ComboBox[,] ComboBoxArrayCircutBreak()
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
        protected ComboBox[,] ComboBoxArraySwitch()
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

        /// <summary>
        /// Функция замены для запроса SQL
        /// </summary>
        /// <returns></returns>
        public static IDictionary<string, string> GetDisconaryVendor()
        {
            Dictionary<string, string> disconaryVendor = new Dictionary<string, string>()
            {
                {"IEK", "Iek"},
                {"IEK BA47", "IekVa47"},
                {"IEK BA47М", "IekVa47m"},
                {"EKF PROxima", "EkfProxima"},
                {"EKF AVERS", "EkfAvers"},
                {"ABB", "Abb"},
                {"KEAZ", "Keaz"},
                {"DKC", "Dkc"},
                {"DEKraft", "Dekraft"},
                {"Schneider", "Schneider"},
                {"TDM", "Tdm"},
                {"IEK Armat", "IekArmat"}
            };
            return disconaryVendor;
        }

        private void MessageError(string textMessage, string textAtribute)
        {
            MessageBox.Show(textMessage,
                textAtribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        #region line1_CircutBreak

        private void checkBox1_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FirstLineArray);


        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FirstLineArray);

        #endregion

        #region line2_CircutBreak

        private void checkBox2_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SecondLineArray);

        #endregion

        #region line3_CircutBreak
        private void checkBox3_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.ThirdLineArray);


        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.ThirdLineArray);

        #endregion

        #region line4_CircutBreak

        private void checkBox4_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox20_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FourthLineArray);

        #endregion

        #region line5_CircutBreak

        private void checkBox5_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox24_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox23_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox22_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox21_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.FifthLineArray);

        #endregion

        #region line6_CircutBreak

        private void checkBox6_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox30_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox29_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox28_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox27_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox26_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircutBreakAsync((int)ContainerAvt.SixthLineArray);

        #endregion

        #region line1_Switch

        private void checkBox7_CheckedChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox35_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox32_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.FirstLineArray);
        private void comboBox31_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.FirstLineArray);

        #endregion

        #region line2_Switch

        private void checkBox8_CheckedChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox40_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox37_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox36_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.SecondLineArray);

        #endregion

        #region line3_Switch

        private void checkBox9_CheckedChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox45_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox42_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox41_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.ThirdLineArray);

        #endregion

        #region line4_Switch

        private void checkBox10_CheckedChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox50_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox47_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox46_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.FourthLineArray);

        #endregion

        #region line5_Switch
        private void checkBox11_CheckedChanged(object sender, EventArgs e) =>
           CheckDataSwitchAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox55_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox52_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox51_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.FifthLineArray);

        #endregion

        #region line6_Switch
        private void checkBox12_CheckedChanged(object sender, EventArgs e) =>
          CheckDataSwitchAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox60_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox57_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox56_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataSwitchAsync((int)ContainerAvt.SixthLineArray);

        #endregion

        #region KeyPress

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace        
                e.Handled = true;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        #endregion

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            var comboBoxes = ComboBoxArrayCircutBreak();
            for (int i = 1; i < 5; i++)
            {
                comboBoxes[1, i].SelectedIndex = comboBoxes[0, i].SelectedIndex;
                comboBoxes[2, i].SelectedIndex = comboBoxes[0, i].SelectedIndex;
                comboBoxes[3, i].SelectedIndex = comboBoxes[0, i].SelectedIndex;
                comboBoxes[4, i].SelectedIndex = comboBoxes[0, i].SelectedIndex;
                comboBoxes[5, i].SelectedIndex = comboBoxes[0, i].SelectedIndex;
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            var comboBoxes = ComboBoxArrayCircutBreak();
            for (int i = 1; i < 5; i++)
            {
                comboBoxes[2, i].SelectedIndex = comboBoxes[1, i].SelectedIndex;
                comboBoxes[3, i].SelectedIndex = comboBoxes[1, i].SelectedIndex;
                comboBoxes[4, i].SelectedIndex = comboBoxes[1, i].SelectedIndex;
                comboBoxes[5, i].SelectedIndex = comboBoxes[1, i].SelectedIndex;
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            var comboBoxes = ComboBoxArrayCircutBreak();
            for (int i = 1; i < 5; i++)
            {
                comboBoxes[3, i].SelectedIndex = comboBoxes[2, i].SelectedIndex;
                comboBoxes[4, i].SelectedIndex = comboBoxes[2, i].SelectedIndex;
                comboBoxes[5, i].SelectedIndex = comboBoxes[2, i].SelectedIndex;
            }
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            var comboBoxes = ComboBoxArrayCircutBreak();
            for (int i = 1; i < 5; i++)
            {
                comboBoxes[4, i].SelectedIndex = comboBoxes[3, i].SelectedIndex;
                comboBoxes[5, i].SelectedIndex = comboBoxes[3, i].SelectedIndex;
            }
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            var comboBoxes = ComboBoxArrayCircutBreak();
            for (int i = 1; i < 5; i++)
            {
                comboBoxes[5, i].SelectedIndex = comboBoxes[4, i].SelectedIndex;
            }
        }
    }
}
