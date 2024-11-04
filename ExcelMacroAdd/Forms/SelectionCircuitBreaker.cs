using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Models;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services.Interfaces;
using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

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

        private readonly IDataInXml dataInXml;
        private readonly ISelectionCircuitBreakerData accessData;
        private UserVariable[] userVariables = new UserVariable[6];

        //Singelton
        private static SelectionCircuitBreaker instance;
        public static void getInstance(IDataInXml dataInXml, ISelectionCircuitBreakerData accessData, IFormSettings formSettings)
        {
            if (instance == null)
            {
                instance = new SelectionCircuitBreaker(dataInXml, accessData)
                {
                    TopMost = formSettings.FormTopMost
                };
                instance.ShowDialog();
            }
        }

        private void SelectionCircuitBreaker_FormClosed(object sender, FormClosedEventArgs e) =>
            instance = null;

        private SelectionCircuitBreaker(IDataInXml dataInXml, ISelectionCircuitBreakerData accessData)
        {
            this.dataInXml = dataInXml;
            this.accessData = accessData;
            InitializeComponent();
        }

        private void SelectionCircuitBreaker_Load(object sender, EventArgs e)
        {

            //Массивы параметров модульных автоматов     
            var loadVendor = accessData.AccessCircuitBreaker.GetAllUniqueVendors();
            ComboBox[] comboBoxItVendor = { comboBox1, comboBox7, comboBox13, comboBox19, comboBox25, comboBox31 };

            for (int i = 0; i < 6; i++)
            {
                //    //Добавление данных по вендорам
                comboBoxItVendor[i].Items.AddRange(loadVendor);
                comboBoxItVendor[i].SelectedIndex = 1;
            }
        }

        #region CheckLine1

        private void checkBox1_CheckedChanged(object sender, EventArgs e) =>
          CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);

        private void textBox1_TextChanged(object sender, EventArgs e) =>
          CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FirstLineArray);

        #endregion

        #region CheckLine2

        private void checkBox2_CheckedChanged(object sender, EventArgs e) =>
             CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        private void textBox2_TextChanged(object sender, EventArgs e) =>
             CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e) =>
             CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e) =>
             CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e) =>
             CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e) =>
             CheckDataCircuitBreakAsync((int)ContainerAvt.SecondLineArray);

        #endregion

        #region CheckLine3

        private void checkBox3_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void textBox3_TextChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.ThirdLineArray);

        #endregion

        #region CheckLine4

        private void checkBox4_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        private void textBox4_TextChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox21_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox22_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox23_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        private void comboBox24_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FourthLineArray);

        #endregion

        #region CheckLine5

        private void checkBox5_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        private void textBox5_TextChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox27_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox28_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox29_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        private void comboBox30_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.FifthLineArray);

        #endregion

        #region CheckLine6

        private void checkBox6_CheckedChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

        private void textBox6_TextChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox33_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox34_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox35_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

        private void comboBox36_SelectedIndexChanged(object sender, EventArgs e) =>
            CheckDataCircuitBreakAsync((int)ContainerAvt.SixthLineArray);

        #endregion

        #region ComboboxSeries
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox1.Text;
            var loadSeries = accessData.AccessCircuitBreaker.GetAllUniqueSeries(vendor);
            comboBox2.Items.Clear();
            comboBox2.Items.AddRange(loadSeries);
            comboBox2.SelectedIndex = 0;
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox7.Text;
            var loadSeries = accessData.AccessCircuitBreaker.GetAllUniqueSeries(vendor);
            comboBox8.Items.Clear();
            comboBox8.Items.AddRange(loadSeries);
            comboBox8.SelectedIndex = 0;
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox13.Text;
            var loadSeries = accessData.AccessCircuitBreaker.GetAllUniqueSeries(vendor);
            comboBox14.Items.Clear();
            comboBox14.Items.AddRange(loadSeries);
            comboBox14.SelectedIndex = 0;
        }

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox19.Text;
            var loadSeries = accessData.AccessCircuitBreaker.GetAllUniqueSeries(vendor);
            comboBox20.Items.Clear();
            comboBox20.Items.AddRange(loadSeries);
            comboBox20.SelectedIndex = 0;
        }

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox25.Text;
            var loadSeries = accessData.AccessCircuitBreaker.GetAllUniqueSeries(vendor);
            comboBox26.Items.Clear();
            comboBox26.Items.AddRange(loadSeries);
            comboBox26.SelectedIndex = 0;
        }

        private void comboBox31_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox31.Text;
            var loadSeries = accessData.AccessCircuitBreaker.GetAllUniqueSeries(vendor);
            comboBox32.Items.Clear();
            comboBox32.Items.AddRange(loadSeries);
            comboBox32.SelectedIndex = 0;
        }
        #endregion

        //Загружаем в Combobox группу, ток, кривую, ток отключения и кол-во полюсов
        #region setAllDataCombobox
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox1.Text;
            string series = comboBox2.Text;
            var data = accessData.AccessCircuitBreaker.GetDataCircutBreaker(vendor, series);
            string group = data.group;
            if (group != null)
            {
                label1.Visible = true;
                label1.Text = group;
            }
            else
            {
                label1.Visible = false;
            }

            comboBox3.Items.Clear();
            comboBox3.Items.AddRange(data.current.Select(i => i.ToString()).ToArray());

            //Костыль
            if (data.current.Count() > 5)
            {
                comboBox3.SelectedIndex = 5;
            }
            else
            {
                comboBox3.SelectedIndex = 0;
            }

            comboBox4.Items.Clear();
            comboBox4.Items.AddRange(data.kurve);
            comboBox4.SelectedIndex = 0;

            comboBox5.Items.Clear();
            comboBox5.Items.AddRange(data.maxCurrent);
            comboBox5.SelectedIndex = 0;

            comboBox6.Items.Clear();
            comboBox6.Items.AddRange(data.quantityPole);
            comboBox6.SelectedIndex = 0;
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox7.Text;
            string series = comboBox8.Text;
            var data = accessData.AccessCircuitBreaker.GetDataCircutBreaker(vendor, series);
            string group = data.group;
            if (group != null)
            {
                label2.Visible = true;
                label2.Text = group;
            }
            else
            {
                label2.Visible = false;
            }

            comboBox9.Items.Clear();
            comboBox9.Items.AddRange(data.current.Select(i => i.ToString()).ToArray());

            //Костыль
            if (data.current.Count() > 5)
            {
                comboBox9.SelectedIndex = 5;
            }
            else
            {
                comboBox9.SelectedIndex = 0;
            }

            comboBox10.Items.Clear();
            comboBox10.Items.AddRange(data.kurve);
            comboBox10.SelectedIndex = 0;

            comboBox11.Items.Clear();
            comboBox11.Items.AddRange(data.maxCurrent);
            comboBox11.SelectedIndex = 0;

            comboBox12.Items.Clear();
            comboBox12.Items.AddRange(data.quantityPole);
            comboBox12.SelectedIndex = 0;
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox13.Text;
            string series = comboBox14.Text;
            var data = accessData.AccessCircuitBreaker.GetDataCircutBreaker(vendor, series);
            string group = data.group;
            if (group != null)
            {
                label3.Visible = true;
                label3.Text = group;
            }
            else
            {
                label3.Visible = false;
            }

            comboBox15.Items.Clear();
            comboBox15.Items.AddRange(data.current.Select(i => i.ToString()).ToArray());

            //Костыль
            if (data.current.Count() > 5)
            {
                comboBox15.SelectedIndex = 5;
            }
            else
            {
                comboBox15.SelectedIndex = 0;
            }

            comboBox16.Items.Clear();
            comboBox16.Items.AddRange(data.kurve);
            comboBox16.SelectedIndex = 0;

            comboBox17.Items.Clear();
            comboBox17.Items.AddRange(data.maxCurrent);
            comboBox17.SelectedIndex = 0;

            comboBox18.Items.Clear();
            comboBox18.Items.AddRange(data.quantityPole);
            comboBox18.SelectedIndex = 0;
        }

        private void comboBox20_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox19.Text;
            string series = comboBox20.Text;
            var data = accessData.AccessCircuitBreaker.GetDataCircutBreaker(vendor, series);
            string group = data.group;
            if (group != null)
            {
                label4.Visible = true;
                label4.Text = group;
            }
            else
            {
                label4.Visible = false;
            }

            comboBox21.Items.Clear();
            comboBox21.Items.AddRange(data.current.Select(i => i.ToString()).ToArray());

            //Костыль
            if (data.current.Count() > 5)
            {
                comboBox21.SelectedIndex = 5;
            }
            else
            {
                comboBox21.SelectedIndex = 0;
            }

            comboBox22.Items.Clear();
            comboBox22.Items.AddRange(data.kurve);
            comboBox22.SelectedIndex = 0;

            comboBox23.Items.Clear();
            comboBox23.Items.AddRange(data.maxCurrent);
            comboBox23.SelectedIndex = 0;

            comboBox24.Items.Clear();
            comboBox24.Items.AddRange(data.quantityPole);
            comboBox24.SelectedIndex = 0;
        }

        private void comboBox26_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox25.Text;
            string series = comboBox26.Text;
            var data = accessData.AccessCircuitBreaker.GetDataCircutBreaker(vendor, series);
            string group = data.group;
            if (group != null)
            {
                label5.Visible = true;
                label5.Text = group;
            }
            else
            {
                label5.Visible = false;
            }

            comboBox27.Items.Clear();
            comboBox27.Items.AddRange(data.current.Select(i => i.ToString()).ToArray());

            //Костыль
            if (data.current.Count() > 5)
            {
                comboBox27.SelectedIndex = 5;
            }
            else
            {
                comboBox27.SelectedIndex = 0;
            }

            comboBox28.Items.Clear();
            comboBox28.Items.AddRange(data.kurve);
            comboBox28.SelectedIndex = 0;

            comboBox29.Items.Clear();
            comboBox29.Items.AddRange(data.maxCurrent);
            comboBox29.SelectedIndex = 0;

            comboBox30.Items.Clear();
            comboBox30.Items.AddRange(data.quantityPole);
            comboBox30.SelectedIndex = 0;
        }

        private void comboBox32_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendor = comboBox31.Text;
            string series = comboBox32.Text;
            var data = accessData.AccessCircuitBreaker.GetDataCircutBreaker(vendor, series);
            string group = data.group;
            if (group != null)
            {
                label6.Visible = true;
                label6.Text = group;
            }
            else
            {
                label6.Visible = false;
            }

            comboBox33.Items.Clear();
            comboBox33.Items.AddRange(data.current.Select(i => i.ToString()).ToArray());

            //Костыль
            if (data.current.Count() > 5)
            {
                comboBox33.SelectedIndex = 5;
            }
            else
            {
                comboBox33.SelectedIndex = 0;
            }

            comboBox34.Items.Clear();
            comboBox34.Items.AddRange(data.kurve);
            comboBox34.SelectedIndex = 0;

            comboBox35.Items.Clear();
            comboBox35.Items.AddRange(data.maxCurrent);
            comboBox35.SelectedIndex = 0;

            comboBox36.Items.Clear();
            comboBox36.Items.AddRange(data.quantityPole);
            comboBox36.SelectedIndex = 0;
        }
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

        #endregion

        private async void CheckDataCircuitBreakAsync(int rowsCheck)
        {

            PictureBox[] pictures = PictureBoxesCircuitBreak();
            CheckBox[] checks = CheckBoxArrayCircuitBreak();

            //Если стоит галочка в CheckBox, то условие истина
            if (!checks[rowsCheck].Checked)
            {
                return;
            }

            var vendor = ComboBoxesArrayVendor()[rowsCheck].SelectedItem.ToString();
            var series = ComboBoxesArraySeries()[rowsCheck].SelectedItem.ToString();
            int.TryParse(ComboBoxesArrayCurrent()[rowsCheck].SelectedItem.ToString(), out int current);
            var kurve = ComboBoxesArrayCurve()[rowsCheck].SelectedItem.ToString();
            var maxCurrent = ComboBoxesArrayMaxCurrent()[rowsCheck].SelectedItem.ToString();
            var polus = ComboBoxesArrayPolus()[rowsCheck].SelectedItem.ToString();
            int.TryParse(TextBoxesArrayCircuitBreak()[rowsCheck].Text, out int quantity);

            try
            {
                var modules = await accessData.AccessCircuitBreaker.GetEntityCircuitBreaker(vendor, series, current, kurve, maxCurrent, polus);

                if (modules != null)
                {
                    UserVariable userVariable = new UserVariable { article = modules.ArticleNumber, vendor = vendor, quantity = quantity, number = rowsCheck };
                    userVariables[rowsCheck] = userVariable;
                    pictures[rowsCheck].BackColor = Color.Green;
                    return;
                }
                userVariables[rowsCheck] = null;
                pictures[rowsCheck].BackColor = Color.IndianRed;
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
            CreateFillInCircutBreakAsync();
        }

        private void CreateFillInCircutBreakAsync()
        {
            int rows = default;
            foreach (var item in userVariables)
            {
                if (item == null) continue;
                if (CheckBoxArrayCircuitBreak()[item.number].Checked)
                {
                    var writeExcel = new WriteExcel(dataInXml, item.vendor, rows++, item.article, item.quantity);
                    writeExcel.Start();
                }
            }
        }

        private PictureBox[] PictureBoxesCircuitBreak() =>
            new PictureBox[] { pictureBox1, pictureBox2, pictureBox3, pictureBox4, pictureBox5, pictureBox6 };

        private TextBox[] TextBoxesArrayCircuitBreak() =>
            new TextBox[] { textBox1, textBox2, textBox3, textBox4, textBox5, textBox6 };

        private CheckBox[] CheckBoxArrayCircuitBreak() =>
            new CheckBox[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 };

        private ComboBox[] ComboBoxesArrayVendor() =>
             new ComboBox[] { comboBox1, comboBox7, comboBox13, comboBox19, comboBox25, comboBox31 };

        private ComboBox[] ComboBoxesArraySeries() =>
             new ComboBox[] { comboBox2, comboBox8, comboBox14, comboBox20, comboBox26, comboBox32 };

        private ComboBox[] ComboBoxesArrayCurrent() =>
            new ComboBox[] { comboBox3, comboBox9, comboBox15, comboBox21, comboBox27, comboBox33 };

        private ComboBox[] ComboBoxesArrayCurve() =>
            new ComboBox[] { comboBox4, comboBox10, comboBox16, comboBox22, comboBox28, comboBox34 };

        private ComboBox[] ComboBoxesArrayMaxCurrent() =>
           new ComboBox[] { comboBox5, comboBox11, comboBox17, comboBox23, comboBox29, comboBox35 };

        private ComboBox[] ComboBoxesArrayPolus() =>
            new ComboBox[] { comboBox6, comboBox12, comboBox18, comboBox24, comboBox30, comboBox36 };


        private static void MessageError(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            var comboBoxesArrayVendor = ComboBoxesArrayVendor();
            var comboBoxesArraySeries = ComboBoxesArraySeries();

            for (int i = 1; i < 6; i++)
            {
                comboBoxesArrayVendor[i].SelectedIndex = comboBoxesArrayVendor[0].SelectedIndex;
                comboBoxesArraySeries[i].SelectedIndex = comboBoxesArraySeries[0].SelectedIndex;
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            var comboBoxesArrayVendor = ComboBoxesArrayVendor();
            var comboBoxesArraySeries = ComboBoxesArraySeries();

            for (int i = 2; i < 6; i++)
            {
                comboBoxesArrayVendor[i].SelectedIndex = comboBoxesArrayVendor[1].SelectedIndex;
                comboBoxesArraySeries[i].SelectedIndex = comboBoxesArraySeries[1].SelectedIndex;
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            var comboBoxesArrayVendor = ComboBoxesArrayVendor();
            var comboBoxesArraySeries = ComboBoxesArraySeries();

            for (int i = 3; i < 6; i++)
            {
                comboBoxesArrayVendor[i].SelectedIndex = comboBoxesArrayVendor[2].SelectedIndex;
                comboBoxesArraySeries[i].SelectedIndex = comboBoxesArraySeries[2].SelectedIndex;
            }
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            var comboBoxesArrayVendor = ComboBoxesArrayVendor();
            var comboBoxesArraySeries = ComboBoxesArraySeries();

            for (int i = 4; i < 6; i++)
            {
                comboBoxesArrayVendor[i].SelectedIndex = comboBoxesArrayVendor[3].SelectedIndex;
                comboBoxesArraySeries[i].SelectedIndex = comboBoxesArraySeries[3].SelectedIndex;
            }
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            var comboBoxesArrayVendor = ComboBoxesArrayVendor();
            var comboBoxesArraySeries = ComboBoxesArraySeries();

            for (int i = 5; i < 6; i++)
            {
                comboBoxesArrayVendor[i].SelectedIndex = comboBoxesArrayVendor[4].SelectedIndex;
                comboBoxesArraySeries[i].SelectedIndex = comboBoxesArraySeries[4].SelectedIndex;
            }
        }
    }
}
