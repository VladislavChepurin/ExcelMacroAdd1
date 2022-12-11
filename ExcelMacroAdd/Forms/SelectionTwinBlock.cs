using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Interfaces;
using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class SelectionTwinBlock : Form
    {
        private const byte StartSwitchCurrent = 5; // Начальный ток трансформации
        private readonly IDataInXml dataInXml;
        private readonly ISelectionTwinBlockData accessData;

        //Singelton
        private static SelectionTwinBlock instance;
        public static async Task getInstance(IDataInXml dataInXml, ISelectionTwinBlockData accessData, IFormSettings formSettings)
        {
            if (instance == null)
            {
                await Task.Run(() =>
                {
                    instance = new SelectionTwinBlock(dataInXml, accessData);
                    instance.TopMost = formSettings.FormTopMost;
                    instance.ShowDialog();
                });
            }
        }

        private void SelectionTwinBlock_FormClosed(object sender, FormClosedEventArgs e) =>      
            instance = null;

        private SelectionTwinBlock(IDataInXml dataInXml, ISelectionTwinBlockData accessData)
        {
            this.dataInXml = dataInXml;
            this.accessData = accessData;
            InitializeComponent();
        }

        private void SelectionTwinBlock_Load(object sender, EventArgs e)
        {
            // ReSharper disable once CoVariantArrayConversion
            comboBox1.Items.AddRange(accessData.AccessTwinBlock.GetComboBox1Items());
            comboBox1.SelectedIndex = StartSwitchCurrent;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Pen pen = new Pen(Color.FromArgb(40, 0, 0, 100));
            e.Graphics.DrawLine(pen, 275, 70, 465, 70);
        }
        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckingAvailableAccessories();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            CheckingAvailableAccessories();
        }

        private void CheckingAvailableAccessories()
        {
            var data = accessData.AccessTwinBlock.GetDataInTableDb(comboBox1.SelectedItem.ToString(), checkBox1.Checked);

            if (!string.IsNullOrEmpty(data.Item1))
            {
                pictureBox1.Image =
                    ByteArrayToImage(accessData.AccessTwinBlock.GetBlobPictureDb(comboBox1.SelectedItem.ToString(),
                        checkBox1.Checked));
                button1.Enabled = true;
            }
            else
            {
                pictureBox1.Image = Properties.Resources.none;
                button1.Enabled = false;
            }

            if (string.IsNullOrEmpty(data.Item2))
            {
                checkBox2.Checked = false;
                checkBox2.Enabled = false;
            }
            else
            {
                checkBox2.Enabled = true;
            }

            if (string.IsNullOrEmpty(data.Item3))
            {
                checkBox3.Checked = false;
                checkBox3.Enabled = false;
            }
            else
            {
                checkBox3.Enabled = true;
            }

            if (string.IsNullOrEmpty(data.Item4))
            {
                checkBox4.Checked = false;
                checkBox4.Enabled = false;
            }
            else
            {
                checkBox4.Enabled = true;
            }

            if (string.IsNullOrEmpty(data.Item5))
            {
                checkBox5.Checked = false;
                checkBox5.Enabled = false;
            }
            else
            {
                checkBox5.Enabled = true;
            }
        }

        private Image ByteArrayToImage(byte[] byteArrayIn)
        {
            using (var ms = new MemoryStream(byteArrayIn))
            {
                var returnImage = Image.FromStream(ms);
                return returnImage;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int rowsLine = default;
            var data = accessData.AccessTwinBlock.GetDataInTableDb(comboBox1.SelectedItem.ToString(), checkBox1.Checked);

            var writeExcel = new WriteExcel(dataInXml, "Ekf", rowsLine, data.Item1);
            writeExcel.Start();

            if (!string.IsNullOrEmpty(data.Item2) && checkBox2.Checked)
            {
                rowsLine++;
                writeExcel = new WriteExcel(dataInXml, "Ekf", rowsLine, data.Item2);
                writeExcel.Start();
            }

            if (!string.IsNullOrEmpty(data.Item3) && checkBox3.Checked)
            {
                rowsLine++;
                writeExcel = new WriteExcel(dataInXml, "Ekf", rowsLine, data.Item3);
                writeExcel.Start();
            }

            if (!string.IsNullOrEmpty(data.Item4) && checkBox4.Checked)
            {
                rowsLine++;
                writeExcel = new WriteExcel(dataInXml, "Ekf", rowsLine, data.Item4);
                writeExcel.Start();
            }

            if (!string.IsNullOrEmpty(data.Item5) && checkBox5.Checked)
            {
                rowsLine++;
                writeExcel = new WriteExcel(dataInXml, "Ekf", rowsLine, data.Item5);
                writeExcel.Start();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
