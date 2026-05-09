using ExcelMacroAdd.BusinessLayer.Interfaces;
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

        private sealed class CircuitBreakerRowControls
        {
            public CircuitBreakerRowControls(
                CheckBox checkBox,
                TextBox quantityTextBox,
                ComboBox vendorComboBox,
                ComboBox seriesComboBox,
                ComboBox currentComboBox,
                ComboBox curveComboBox,
                ComboBox maxCurrentComboBox,
                ComboBox polesComboBox,
                PictureBox statusPictureBox,
                Label groupLabel)
            {
                CheckBox = checkBox;
                QuantityTextBox = quantityTextBox;
                VendorComboBox = vendorComboBox;
                SeriesComboBox = seriesComboBox;
                CurrentComboBox = currentComboBox;
                CurveComboBox = curveComboBox;
                MaxCurrentComboBox = maxCurrentComboBox;
                PolesComboBox = polesComboBox;
                StatusPictureBox = statusPictureBox;
                GroupLabel = groupLabel;
            }

            public CheckBox CheckBox { get; }

            public TextBox QuantityTextBox { get; }

            public ComboBox VendorComboBox { get; }

            public ComboBox SeriesComboBox { get; }

            public ComboBox CurrentComboBox { get; }

            public ComboBox CurveComboBox { get; }

            public ComboBox MaxCurrentComboBox { get; }

            public ComboBox PolesComboBox { get; }

            public PictureBox StatusPictureBox { get; }

            public Label GroupLabel { get; }
        }

        private readonly IDataInXml dataInXml;
        private readonly ISelectionCircuitBreakerData accessData;
        private readonly CircuitBreakerRowControls[] circuitBreakerRows;
        private UserVariable[] userVariables = new UserVariable[6];

        private void SelectionCircuitBreaker_FormClosed(object sender, FormClosedEventArgs e)
        {
            SelectionModularDevices main = this.Owner as SelectionModularDevices;
            main?.Show();
        }

        public SelectionCircuitBreaker(IDataInXml dataInXml, ISelectionCircuitBreakerData accessData, IFormSettings formSettings)
        {
            TopMost = formSettings.FormTopMost;
            this.dataInXml = dataInXml;
            this.accessData = accessData;
            InitializeComponent();
            circuitBreakerRows = CreateRows();
        }

        private void SelectionCircuitBreaker_Load(object sender, EventArgs e)
        {
            var loadVendor = accessData.AccessCircuitBreaker.GetAllUniqueVendors();

            foreach (var row in circuitBreakerRows)
            {
                row.VendorComboBox.Items.AddRange(loadVendor);
                row.VendorComboBox.SelectedIndex = 1;
            }
        }

        #region CheckLine1

        private void checkBox1_CheckedChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FirstLineArray);

        private void textBox1_TextChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FirstLineArray);

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FirstLineArray);

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FirstLineArray);

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FirstLineArray);

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FirstLineArray);

        #endregion

        #region CheckLine2

        private void checkBox2_CheckedChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SecondLineArray);

        private void textBox2_TextChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SecondLineArray);

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SecondLineArray);

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SecondLineArray);

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SecondLineArray);

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SecondLineArray);

        #endregion

        #region CheckLine3

        private void checkBox3_CheckedChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.ThirdLineArray);

        private void textBox3_TextChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.ThirdLineArray);

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.ThirdLineArray);

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.ThirdLineArray);

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.ThirdLineArray);

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.ThirdLineArray);

        #endregion

        #region CheckLine4

        private void checkBox4_CheckedChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FourthLineArray);

        private void textBox4_TextChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FourthLineArray);

        private void comboBox21_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FourthLineArray);

        private void comboBox22_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FourthLineArray);

        private void comboBox23_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FourthLineArray);

        private void comboBox24_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FourthLineArray);

        #endregion

        #region CheckLine5

        private void checkBox5_CheckedChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FifthLineArray);

        private void textBox5_TextChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FifthLineArray);

        private void comboBox27_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FifthLineArray);

        private void comboBox28_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FifthLineArray);

        private void comboBox29_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FifthLineArray);

        private void comboBox30_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.FifthLineArray);

        #endregion

        #region CheckLine6

        private void checkBox6_CheckedChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SixthLineArray);

        private void textBox6_TextChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SixthLineArray);

        private void comboBox33_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SixthLineArray);

        private void comboBox34_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SixthLineArray);

        private void comboBox35_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SixthLineArray);

        private void comboBox36_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleRowSelectionChanged((int)ContainerAvt.SixthLineArray);

        #endregion

        #region ComboboxSeries

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleVendorChanged((int)ContainerAvt.FirstLineArray);

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleVendorChanged((int)ContainerAvt.SecondLineArray);

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleVendorChanged((int)ContainerAvt.ThirdLineArray);

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleVendorChanged((int)ContainerAvt.FourthLineArray);

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleVendorChanged((int)ContainerAvt.FifthLineArray);

        private void comboBox31_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleVendorChanged((int)ContainerAvt.SixthLineArray);

        #endregion

        #region setAllDataCombobox

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleSeriesChanged((int)ContainerAvt.FirstLineArray);

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleSeriesChanged((int)ContainerAvt.SecondLineArray);

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleSeriesChanged((int)ContainerAvt.ThirdLineArray);

        private void comboBox20_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleSeriesChanged((int)ContainerAvt.FourthLineArray);

        private void comboBox26_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleSeriesChanged((int)ContainerAvt.FifthLineArray);

        private void comboBox32_SelectedIndexChanged(object sender, EventArgs e) =>
            HandleSeriesChanged((int)ContainerAvt.SixthLineArray);

        #endregion

        #region KeyPress

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e) =>
            HandleQuantityKeyPress(e);

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e) =>
            HandleQuantityKeyPress(e);

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e) =>
            HandleQuantityKeyPress(e);

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e) =>
            HandleQuantityKeyPress(e);

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e) =>
            HandleQuantityKeyPress(e);

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e) =>
            HandleQuantityKeyPress(e);

        #endregion

        private void HandleRowSelectionChanged(int rowIndex)
        {
            CheckDataCircuitBreakAsync(rowIndex);
        }

        private void HandleVendorChanged(int rowIndex)
        {
            var row = circuitBreakerRows[rowIndex];
            string vendor = row.VendorComboBox.Text;
            var loadSeries = accessData.AccessCircuitBreaker.GetAllUniqueSeries(vendor);

            row.SeriesComboBox.Items.Clear();
            row.SeriesComboBox.Items.AddRange(loadSeries);
            row.SeriesComboBox.SelectedIndex = 0;
        }

        private void HandleSeriesChanged(int rowIndex)
        {
            var row = circuitBreakerRows[rowIndex];
            string vendor = row.VendorComboBox.Text;
            string series = row.SeriesComboBox.Text;
            var data = accessData.AccessCircuitBreaker.GetDataCircutBreaker(vendor, series);

            SetGroupLabel(row.GroupLabel, data.group);

            row.CurrentComboBox.Items.Clear();
            row.CurrentComboBox.Items.AddRange(data.current.Select(i => i.ToString()).ToArray());
            row.CurrentComboBox.SelectedIndex = data.current.Count() > 5 ? 5 : 0;

            row.CurveComboBox.Items.Clear();
            row.CurveComboBox.Items.AddRange(data.kurve);
            row.CurveComboBox.SelectedIndex = 0;

            row.MaxCurrentComboBox.Items.Clear();
            row.MaxCurrentComboBox.Items.AddRange(data.maxCurrent);
            row.MaxCurrentComboBox.SelectedIndex = 0;

            row.PolesComboBox.Items.Clear();
            row.PolesComboBox.Items.AddRange(data.quantityPole);
            row.PolesComboBox.SelectedIndex = 0;
        }

        private static void SetGroupLabel(Label label, string group)
        {
            if (group != null)
            {
                label.Visible = true;
                label.Text = group;
                return;
            }

            label.Visible = false;
        }

        private static void HandleQuantityKeyPress(KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!char.IsDigit(number) && number != 8)
            {
                e.Handled = true;
            }
        }

        private async void CheckDataCircuitBreakAsync(int rowsCheck)
        {
            var row = circuitBreakerRows[rowsCheck];

            if (!row.CheckBox.Checked)
            {
                return;
            }

            var vendor = row.VendorComboBox.SelectedItem.ToString();
            var series = row.SeriesComboBox.SelectedItem.ToString();
            int.TryParse(row.CurrentComboBox.SelectedItem.ToString(), out int current);
            var curve = row.CurveComboBox.SelectedItem.ToString();
            var maxCurrent = row.MaxCurrentComboBox.SelectedItem.ToString();
            var poles = row.PolesComboBox.SelectedItem.ToString();
            int.TryParse(row.QuantityTextBox.Text, out int quantity);

            try
            {
                var modules = await accessData.AccessCircuitBreaker.GetEntityCircuitBreaker(
                    vendor,
                    series,
                    current,
                    curve,
                    maxCurrent,
                    poles);

                if (modules != null)
                {
                    UserVariable userVariable = new UserVariable
                    {
                        article = modules.ArticleNumber,
                        vendor = vendor,
                        quantity = quantity,
                        number = rowsCheck
                    };

                    userVariables[rowsCheck] = userVariable;
                    row.StatusPictureBox.BackColor = Color.Green;
                    return;
                }

                userVariables[rowsCheck] = null;
                row.StatusPictureBox.BackColor = Color.IndianRed;
            }
            catch (DataException)
            {
                MessageError(
                    "Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                    "Ошибка базы данных");
            }
            catch (Exception e)
            {
                MessageError(
                    $"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {e.Message}",
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
            using (var scope = new ExcelPerformanceScope(Globals.ThisAddIn.GetApplication()))
            {
                int offsetRow = default;
                foreach (var item in userVariables)
                {
                    if (item == null)
                    {
                        continue;
                    }

                    if (circuitBreakerRows[item.number].CheckBox.Checked)
                    {
                        var writeExcel = new WriteExcel(dataInXml, item.vendor, item.article, offsetRow++, item.quantity);
                        writeExcel.Start();
                    }
                }
            }
        }

        private CircuitBreakerRowControls[] CreateRows() =>
            new CircuitBreakerRowControls[]
            {
                new CircuitBreakerRowControls(checkBox1, textBox1, comboBox1, comboBox2, comboBox3, comboBox4, comboBox5, comboBox6, pictureBox1, label1),
                new CircuitBreakerRowControls(checkBox2, textBox2, comboBox7, comboBox8, comboBox9, comboBox10, comboBox11, comboBox12, pictureBox2, label2),
                new CircuitBreakerRowControls(checkBox3, textBox3, comboBox13, comboBox14, comboBox15, comboBox16, comboBox17, comboBox18, pictureBox3, label3),
                new CircuitBreakerRowControls(checkBox4, textBox4, comboBox19, comboBox20, comboBox21, comboBox22, comboBox23, comboBox24, pictureBox4, label4),
                new CircuitBreakerRowControls(checkBox5, textBox5, comboBox25, comboBox26, comboBox27, comboBox28, comboBox29, comboBox30, pictureBox5, label5),
                new CircuitBreakerRowControls(checkBox6, textBox6, comboBox31, comboBox32, comboBox33, comboBox34, comboBox35, comboBox36, pictureBox6, label6)
            };

        private static void MessageError(string textMessage, string textAttribute)
        {
            MessageBox.Show(
                textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        private void pictureBox1_Click(object sender, EventArgs e) =>
            CopyVendorAndSeriesDown((int)ContainerAvt.FirstLineArray);

        private void pictureBox2_Click(object sender, EventArgs e) =>
            CopyVendorAndSeriesDown((int)ContainerAvt.SecondLineArray);

        private void pictureBox3_Click(object sender, EventArgs e) =>
            CopyVendorAndSeriesDown((int)ContainerAvt.ThirdLineArray);

        private void pictureBox4_Click(object sender, EventArgs e) =>
            CopyVendorAndSeriesDown((int)ContainerAvt.FourthLineArray);

        private void pictureBox5_Click(object sender, EventArgs e) =>
            CopyVendorAndSeriesDown((int)ContainerAvt.FifthLineArray);

        private void CopyVendorAndSeriesDown(int sourceRowIndex)
        {
            for (int i = sourceRowIndex + 1; i < circuitBreakerRows.Length; i++)
            {
                circuitBreakerRows[i].VendorComboBox.SelectedIndex = circuitBreakerRows[sourceRowIndex].VendorComboBox.SelectedIndex;
                circuitBreakerRows[i].SeriesComboBox.SelectedIndex = circuitBreakerRows[sourceRowIndex].SeriesComboBox.SelectedIndex;
            }
        }
    }
}
