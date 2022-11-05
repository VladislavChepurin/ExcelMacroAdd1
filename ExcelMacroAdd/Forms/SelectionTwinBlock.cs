using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class SelectionTwinBlock : Form
    {
        private const byte StartSwitchCurrent = 5; // Начальный ток трансформации
        private readonly IDataInXml dataInXml;
        private readonly ISelectionTwinBlockData accessData;

        public SelectionTwinBlock(IDataInXml dataInXml, ISelectionTwinBlockData accessData)
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
            //Код для определения доступных аксесуаров
            //Что недоступно - неактивно
            var checkBoxs = new CheckBox[] { checkBox2, checkBox3, checkBox4, checkBox5 };
            
            var items = new Queue<string>(4);
            items.Enqueue(data.Item2);
            items.Enqueue(data.Item3);
            items.Enqueue(data.Item4);
            items.Enqueue(data.Item5);

            foreach (var iCheckBox in checkBoxs)
            {
                if (string.IsNullOrEmpty(items.Dequeue()))
                {
                    iCheckBox.Checked = false;
                    iCheckBox.Enabled = false;
                }
                else
                {
                    iCheckBox.Enabled = true;
                }
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
