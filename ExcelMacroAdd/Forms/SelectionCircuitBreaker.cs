using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Switch = ExcelMacroAdd.DataLayer.Entity.Switch;

namespace ExcelMacroAdd.Forms
{

    internal partial class SelectionCircuitBreaker : Form
    {
        private enum ContainerAvt
        {
            FirstLineArray,
            SecondLineArray,
            ThirdLineArray,
            FourthLineArray,
            FifthLineArray,
            SixthLineArray
        }

        private static int CircuitIndAvt; // Начальный ток автоматических выключателей
        private static int CurveIndAvt;// Начальная кривая автоматических выключателей
        private static int IcuIndAvt; // Начальная отключающая способность автоматических выключателей
        private static int PolusIndAvt; // Начальная кол-во полюсов автоматических выключателей
        private static int VendorIndAvt; // Начальный вендор автолмтических выключателей

        private static int CircuitIndVn; // Начальный ток выключателей нагрузки
        private static int PolusIndVn; // Начальная кол-во полюсов выключателей нагрузки
        private static int VendorIndVn; // Начальный вендор выключателей нагрузки

        private readonly IDataInXml dataInXml;
        private readonly ISelectionCircuitBreakerData accessData;

        static SelectionCircuitBreaker()
        {
            CircuitIndAvt = 5;
            CurveIndAvt = 1;
            IcuIndAvt = 0;
            PolusIndAvt = 0;
            VendorIndAvt = 0;

            CircuitIndVn = 0;
            PolusIndVn = 0;
            VendorIndVn = 0;
        }

        //Singelton
        private static SelectionCircuitBreaker instance;
        public static async Task getInstance(IDataInXml dataInXml, ISelectionCircuitBreakerData accessData, IFormSettings formSettings)
        {
            if (instance == null)
            {
                await Task.Run(() =>
                {
                    instance = new SelectionCircuitBreaker(dataInXml, accessData);
                    instance.TopMost = formSettings.FormTopMost;
                    instance.ShowDialog();
                });
            }    
        }

        private void SelectionCircuitBreaker_FormClosed(object sender, FormClosedEventArgs e)=>
            instance = null;

        private void SelectionCircuitBreaker_FormClosing(object sender, FormClosingEventArgs e)
        {
            CurveIndAvt = comboBox4.SelectedIndex;
            IcuIndAvt = comboBox3.SelectedIndex;
            PolusIndAvt = comboBox2.SelectedIndex;
            VendorIndAvt = comboBox1.SelectedIndex;

            PolusIndVn = comboBox32.SelectedIndex;
            VendorIndVn = comboBox31.SelectedIndex;
        }

        private SelectionCircuitBreaker(IDataInXml dataInXml, ISelectionCircuitBreakerData accessData)
        {      
            this.dataInXml = dataInXml;
            this.accessData = accessData;
            InitializeComponent();
        }

        private void SelectionCircuitBreaker_Load(object sender, EventArgs e)
        {
            //Массивы параметров модульных автоматов
            var circuitBreakerCurrent = accessData.AccessCircuitBreaker.GetCircuitCurrentItems();
            var circuitBreakerCurve = accessData.AccessCircuitBreaker.GetCircuitCurveItems();
            var maxCircuitBreakerCurrent = accessData.AccessCircuitBreaker.GetCircuitMaxCurrentItems();
            var amountOfPolesCircuitBreaker = accessData.AccessCircuitBreaker.GetCircuitPolesItems();
            var circuitBreakerVendor = new[] { "IEK BA47", "IEK BA47М", "IEK Armat", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DKC", "DEKraft", "Schneider", "TDM" };
            //Массивы параметров выключателей нагрузки
            var loadSwitchCurrent = accessData.AccessCircuitBreaker.GetCircuitSwitchsItems();
            var amountOfPolesLoadSwitch = accessData.AccessCircuitBreaker.GetSwitchsPolesItems();
            var loadSwitchVendor = new[] { "IEK", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DEKraft", "Schneider", "TDM" };

            //Создание массивов ComboBox для автоматических выключателей
            ComboBox[] comboBoxItCircuit = { comboBox5, comboBox10, comboBox15, comboBox20, comboBox25, comboBox30 };
            ComboBox[] comboBoxItCurve = { comboBox4, comboBox9, comboBox14, comboBox19, comboBox24, comboBox29 };
            ComboBox[] comboBoxItIcu = { comboBox3, comboBox8, comboBox13, comboBox18, comboBox23, comboBox28 };
            ComboBox[] comboBoxItPolus = { comboBox2, comboBox7, comboBox12, comboBox17, comboBox22, comboBox27 };
            ComboBox[] comboBoxItVendor = { comboBox1, comboBox6, comboBox11, comboBox16, comboBox21, comboBox26 };
            //Создание массивов ComboBox для выключателей нагрузки
            ComboBox[] comboBoxItCircuitVn = { comboBox35, comboBox40, comboBox45, comboBox50, comboBox55, comboBox60 };
            ComboBox[] comboBoxItPolusVn = { comboBox32, comboBox37, comboBox42, comboBox47, comboBox52, comboBox57 };
            ComboBox[] comboBoxItVendorVn = { comboBox31, comboBox36, comboBox41, comboBox46, comboBox51, comboBox56 };

            for (int i = 0; i < 6; i++)
            {
                //Добавление в модульные автоматы данных тока
                // ReSharper disable once CoVariantArrayConversion
                comboBoxItCircuit[i].Items.AddRange(circuitBreakerCurrent);
                comboBoxItCircuit[i].SelectedIndex = CircuitIndAvt;
                //Добавление в модульные автоматы данных по кривой
                // ReSharper disable once CoVariantArrayConversion
                comboBoxItCurve[i].Items.AddRange(circuitBreakerCurve);
                comboBoxItCurve[i].SelectedIndex = CurveIndAvt;
                //Добавление в модульные автоматы данных по макс току
                // ReSharper disable once CoVariantArrayConversion
                comboBoxItIcu[i].Items.AddRange(maxCircuitBreakerCurrent);
                comboBoxItIcu[i].SelectedIndex = IcuIndAvt;
                //Добавление в модульные автоматы данных по полюсам
                // ReSharper disable once CoVariantArrayConversion
                comboBoxItPolus[i].Items.AddRange(amountOfPolesCircuitBreaker);
                comboBoxItPolus[i].SelectedIndex = PolusIndAvt;
                //Добавление в модульные автоматы данных по вендорам
                // ReSharper disable once CoVariantArrayConversion
                comboBoxItVendor[i].Items.AddRange(circuitBreakerVendor);
                comboBoxItVendor[i].SelectedIndex = VendorIndAvt;
                //Добавление в выключатели нагрузки данных тока
                // ReSharper disable once CoVariantArrayConversion
                comboBoxItCircuitVn[i].Items.AddRange(loadSwitchCurrent);
                comboBoxItCircuitVn[i].SelectedIndex = CircuitIndVn;
                //Добавление в выключатели нагрузки данных по полюсам
                // ReSharper disable once CoVariantArrayConversion
                comboBoxItPolusVn[i].Items.AddRange(amountOfPolesLoadSwitch);
                comboBoxItPolusVn[i].SelectedIndex = PolusIndVn;
                //Добавление в выключатели нагрузки данных по вендорам
                // ReSharper disable once CoVariantArrayConversion
                comboBoxItVendorVn[i].Items.AddRange(loadSwitchVendor);
                comboBoxItVendorVn[i].SelectedIndex = VendorIndVn;
            }
        }

        private async void CheckDataCircuitBreakAsync(int rowsCheck)
        {
            PictureBox[] pictures = PictureBoxesCircuitBreak();
            CheckBox[] checks = CheckBoxArrayCircuitBreak();
            ComboBox[,] comboBoxes = ComboBoxArrayCircuitBreak();
            // Если стоит галочка в CheckBox, то условие истина
            if (!checks[rowsCheck].Checked)
            {
                return;
            }

            var current = comboBoxes[rowsCheck, 0].SelectedItem.ToString();
            var curve = comboBoxes[rowsCheck, 1].SelectedItem.ToString();
            var maxCurrent = comboBoxes[rowsCheck, 2].SelectedItem.ToString();
            var polus = comboBoxes[rowsCheck, 3].SelectedItem.ToString();
            var vendor = GetDictionaryVendor()[comboBoxes[rowsCheck, 4].SelectedItem.ToString()];

            try
            {
                var  modules = await accessData.AccessCircuitBreaker.GetEntityModule(current, curve, maxCurrent, polus);        

                if (modules is null)
                {
                    pictures[rowsCheck].BackColor = Color.IndianRed;
                    return;
                }

                Type myType = typeof(Modul);
                // получаем свойство
                var articleProp = myType.GetProperty(vendor);
                // получаем значение свойства
                var article = articleProp?.GetValue(modules);

                pictures[rowsCheck].BackColor = string.IsNullOrEmpty(article?.ToString()) ? Color.IndianRed : Color.Green;
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

            var current = comboBoxes[rowsCheck, 0].SelectedItem.ToString();
            var polus = comboBoxes[rowsCheck, 1].SelectedItem.ToString();
            var vendor = GetDictionaryVendor()[comboBoxes[rowsCheck, 2].SelectedItem.ToString()];

            try
            {
                var switches = await accessData.AccessCircuitBreaker.GetEntitySwitch(current, polus);             

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

                pictures[rowsCheck].BackColor = string.IsNullOrEmpty(article?.ToString()) ? Color.IndianRed : Color.Green;
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

        private async void CreateFillInCircutBreakAsync()
        {
            CheckBox[] checks = CheckBoxArrayCircuitBreak();
            ComboBox[,] comboBoxes = ComboBoxArrayCircuitBreak();
            TextBox[] texts = TextBoxesArrayCircuitBreak();

            for (int rows = 0; rows < 6; rows++)
            {
                // Если стоит галочка в CheckBox, то условие истина
                if (checks[rows].Checked)
                {
                    var current = comboBoxes[rows, 0].SelectedItem.ToString();
                    var curve = comboBoxes[rows, 1].SelectedItem.ToString();
                    var maxCurrent = comboBoxes[rows, 2].SelectedItem.ToString();
                    var polus = comboBoxes[rows, 3].SelectedItem.ToString();
                    var vendor = GetDictionaryVendor()[comboBoxes[rows, 4].SelectedItem.ToString()];

                    try
                    {
                        var modules = await accessData.AccessCircuitBreaker.GetEntityModule(current, curve, maxCurrent, polus);

                        Type myType = typeof(Modul);
                        // получаем свойство
                        var articleProp = myType.GetProperty(vendor);
                        // получаем значение свойства
                        var article = articleProp?.GetValue(modules);

                        if (string.IsNullOrEmpty(article?.ToString()))
                        {
                            continue;
                        }

                        int.TryParse(texts[rows].Text, out int quantity);
                        var writeExcel = new WriteExcel(dataInXml, vendor, rows, article.ToString(), quantity);
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

        private async void CreateFillInSwitchAsync()
        {
            CheckBox[] checks = CheckBoxArraySwitch();
            ComboBox[,] comboBoxes = ComboBoxArraySwitch();
            TextBox[] texts = TextBoxesArraySwitch();

            for (int rows = 0; rows < 6; rows++)
            {
                // Если стоит галочка в CheckBox, то условие истина
                if (checks[rows].Checked)
                {
                    var current = comboBoxes[rows, 0].SelectedItem.ToString();
                    var polus = comboBoxes[rows, 1].SelectedItem.ToString();
                    var vendor = GetDictionaryVendor()[comboBoxes[rows, 2].SelectedItem.ToString()];

                    try
                    {
                        var switches = await accessData.AccessCircuitBreaker.GetEntitySwitch(current, polus);

                        Type myType = typeof(Switch);
                        // получаем свойство
                        var articleProp = myType.GetProperty(vendor);
                        // получаем значение свойства
                        var article = articleProp?.GetValue(switches);

                        if (string.IsNullOrEmpty(article?.ToString()))
                        {
                            continue;
                        }

                        int.TryParse(texts[rows].Text, out int quantity);
                        var writeExcel = new WriteExcel(dataInXml, vendor, rows, article.ToString(), quantity);
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

        private PictureBox[] PictureBoxesCircuitBreak()
        {
            PictureBox[] pictures = { pictureBox1, pictureBox2, pictureBox3, pictureBox4, pictureBox5, pictureBox6 };
            return pictures;
        }
        private PictureBox[] PictureBoxesSwitch()
        {
            PictureBox[] pictures = { pictureBox7, pictureBox8, pictureBox9, pictureBox10, pictureBox11, pictureBox12 };
            return pictures;
        }
        private TextBox[] TextBoxesArrayCircuitBreak()
        {
            TextBox[] texts = { textBox1, textBox2, textBox3, textBox4, textBox5, textBox6 };
            return texts;
        }
        private TextBox[] TextBoxesArraySwitch()
        {
            TextBox[] texts = { textBox7, textBox8, textBox9, textBox10, textBox11, textBox12 };
            return texts;
        }
        private CheckBox[] CheckBoxArrayCircuitBreak()
        {
            CheckBox[] checks = { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 };
            return checks;
        }
        private CheckBox[] CheckBoxArraySwitch()
        {
            CheckBox[] checks = { checkBox7, checkBox8, checkBox9, checkBox10, checkBox11, checkBox12 };
            return checks;
        }
        private ComboBox[,] ComboBoxArrayCircuitBreak()
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

        /// <summary>
        /// Функция замены для запроса SQL
        /// </summary>
        /// <returns></returns>
        private static IDictionary<string, string> GetDictionaryVendor()
        {
            Dictionary<string, string> dictionaryVendor = new Dictionary<string, string>()
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
            return dictionaryVendor;
        }

        private static void MessageError(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        #region line1_CircutBreak

        private void checkBox1_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);


        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);

        #endregion

        #region line2_CircutBreak

        private void checkBox2_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        #endregion

        #region line3_CircutBreak
        private void checkBox3_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);


        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);

        #endregion

        #region line4_CircutBreak

        private void checkBox4_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox20_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        #endregion

        #region line5_CircutBreak

        private void checkBox5_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox24_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox23_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox22_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox21_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        #endregion

        #region line6_CircutBreak

        private void checkBox6_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox30_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox29_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox28_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox27_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox26_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

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
            var comboBoxes = ComboBoxArrayCircuitBreak();
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
            var comboBoxes = ComboBoxArrayCircuitBreak();
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
            var comboBoxes = ComboBoxArrayCircuitBreak();
            for (int i = 1; i < 5; i++)
            {
                comboBoxes[3, i].SelectedIndex = comboBoxes[2, i].SelectedIndex;
                comboBoxes[4, i].SelectedIndex = comboBoxes[2, i].SelectedIndex;
                comboBoxes[5, i].SelectedIndex = comboBoxes[2, i].SelectedIndex;
            }
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            var comboBoxes = ComboBoxArrayCircuitBreak();
            for (int i = 1; i < 5; i++)
            {
                comboBoxes[4, i].SelectedIndex = comboBoxes[3, i].SelectedIndex;
                comboBoxes[5, i].SelectedIndex = comboBoxes[3, i].SelectedIndex;
            }
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            var comboBoxes = ComboBoxArrayCircuitBreak();
            for (int i = 1; i < 5; i++)
            {
                comboBoxes[5, i].SelectedIndex = comboBoxes[4, i].SelectedIndex;
            }
        }       
    }
}
