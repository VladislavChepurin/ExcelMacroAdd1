using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class SelectionTwinBlock : Form
    {
        private const byte StartSwitchCurrent = 5;
        private readonly IDataInXml dataInXml;
        private readonly ITwinBlockService twinBlockService;

        public SelectionTwinBlock(IDataInXml dataInXml, ITwinBlockService twinBlockService, IFormSettings formSettings)
        {
            InitializeComponent();
            TopMost = formSettings.FormTopMost;
            this.dataInXml = dataInXml;
            this.twinBlockService = twinBlockService;

            comboBoxCurrent.SelectedIndexChanged += (s, e) => CheckingAvailableAccessories();
            checkBoxReverse.CheckedChanged += (s, e) => CheckingAvailableAccessories();
        }

        private void SelectionTwinBlock_Load(object sender, EventArgs e)
        {
            comboBoxCurrent.Items.AddRange(twinBlockService.GetComboBox1Items());
            comboBoxCurrent.SelectedIndex = StartSwitchCurrent;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            using (var pen = new Pen(Color.FromArgb(40, 0, 0, 100)))
            {
                e.Graphics.DrawLine(pen, 275, 70, 465, 70);
            }
        }

        private void CheckingAvailableAccessories()
        {
            var data = twinBlockService.GetDataInTableDb(comboBoxCurrent.SelectedItem.ToString(), checkBoxReverse.Checked);

            if (!string.IsNullOrEmpty(data.Item1))
            {
                pictureBox.Image =
                    ByteArrayToImage(twinBlockService.GetBlobPictureDb(comboBoxCurrent.SelectedItem.ToString(),
                        checkBoxReverse.Checked));
                btnGoSheet.Enabled = true;
            }
            else
            {
                pictureBox.Image = Properties.Resources.none;
                btnGoSheet.Enabled = false;
            }

            if (string.IsNullOrEmpty(data.Item2))
            {
                checkBoxDirectMountingHandle.Checked = false;
                checkBoxDirectMountingHandle.Enabled = false;
            }
            else
            {
                checkBoxDirectMountingHandle.Enabled = true;
            }

            UpdateCheckBox(checkBoxDirectMountingHandle, data.Item2);
            UpdateCheckBox(checkBoxHandleOnDoor, data.Item3);
            UpdateCheckBox(checkBoxHandleRod, data.Item4);
            UpdateCheckBox(checkBoxAdditionalPole, data.Item5);

        }

        private void UpdateCheckBox(CheckBox checkBox, string itemValue)
        {
            bool shouldEnable = !string.IsNullOrEmpty(itemValue);
            checkBox.Enabled = shouldEnable;
            if (!shouldEnable) checkBox.Checked = false;
        }

        private Image ByteArrayToImage(byte[] byteArrayIn)
        {
            using (var ms = new MemoryStream(byteArrayIn))
            {
                var returnImage = Image.FromStream(ms);
                return returnImage;
            }
        }

        private void btnGoSheet_Click(object sender, EventArgs e)
        {
            using (var scope = new ExcelPerformanceScope(Globals.ThisAddIn.GetApplication()))
            {
                int offsetRow = 0;
                var data = twinBlockService.GetDataInTableDb(
                    comboBoxCurrent.SelectedItem.ToString(), checkBoxReverse.Checked);

                var writeExcel = new WriteExcel(dataInXml, "EKF", data.Item1, offsetRow);
                writeExcel.Start();

                var itemsToProcess = new List<(string ItemValue, CheckBox CheckBox)>
            {
                (data.Item2, checkBoxDirectMountingHandle),
                (data.Item3, checkBoxHandleOnDoor),
                (data.Item4, checkBoxHandleRod),
                (data.Item5, checkBoxAdditionalPole)
            };

                foreach (var item in itemsToProcess)
                {
                    if (item.CheckBox.Checked && !string.IsNullOrEmpty(item.ItemValue))
                    {
                        writeExcel = new WriteExcel(dataInXml, "EKF", item.ItemValue, ++offsetRow);
                        writeExcel.Start();
                    }
                }
            }
        }
    }
}
