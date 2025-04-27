using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Models;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class AdditionalDevicesForm : Form
    {
        private readonly IDataInXml dataInXml;
        private readonly IAccessAdditionalModularDevicesData accessData;
        protected readonly Worksheet Worksheet = Globals.ThisAddIn.GetActiveWorksheet();
        protected readonly Range Cell = Globals.ThisAddIn.GetActiveCell();
        private readonly int currentRow;
        private readonly AdditionalDevices circuitBreakerData;
        private readonly AdditionalDevices switchData;
        private readonly AdditionalDevices addDevicesAgregate;
        private bool AllNull(params string[] values) => values.All(v => v == null);

        private void AdditionalDevicesForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            SelectionModularDevices main = this.Owner as SelectionModularDevices;
            main?.Show();
        }

        public AdditionalDevicesForm(IDataInXml dataInXml, IAccessAdditionalModularDevicesData accessData, IFormSettings formSettings)
        {
            TopMost = formSettings.FormTopMost;
            this.dataInXml = dataInXml;
            this.accessData = accessData;
            InitializeComponent();

            currentRow = Cell.Row;
            string sArticle = Convert.ToString(Worksheet.Cells[currentRow, 1].Value2);

            if (string.IsNullOrEmpty(sArticle))
            {
                MessageBox.Show("Ошибка: не указан артикул.");
                Close();
            }

            circuitBreakerData = accessData.AccessAdditionalModularDevices.GetEntityAdditionalCircuitBreaker(sArticle);
            switchData = accessData.AccessAdditionalModularDevices.GetEntityAdditionalSwitch(sArticle);

            if (circuitBreakerData.vendor != null)
            {
                addDevicesAgregate = circuitBreakerData;
            }
            else
            {
                addDevicesAgregate = switchData;
            }
        }      

        private void AdditionalDevicesForm_Load(object sender, EventArgs e)
        {
            if (circuitBreakerData == null && switchData == null)
                return;

            bool additionalDevicesAgregateCheck = AllNull(
                addDevicesAgregate.shuntTrip24vArticle,
                addDevicesAgregate.shuntTrip48vArticle,
                addDevicesAgregate.shuntTrip230vArticle,
                addDevicesAgregate.undervoltageReleaseArticle,
                addDevicesAgregate.signalContactArticle,
                addDevicesAgregate.auxiliaryContactArticle,
                addDevicesAgregate.signalOrAuxiliaryContactArticle);          

            if (!additionalDevicesAgregateCheck)
            {
                btnApply.Enabled = true;
                label1.Visible = false;
                if (addDevicesAgregate.shuntTrip24vArticle != null)
                    checkBoxShuntTrip24V.Enabled = true;
                if (addDevicesAgregate.shuntTrip48vArticle != null)
                    checkBoxShuntTrip48V.Enabled = true;
                if (addDevicesAgregate.shuntTrip230vArticle != null)
                    checkBoxShuntTrip230V.Enabled = true;
                if (addDevicesAgregate.undervoltageReleaseArticle != null)
                    checkBoxUndervoltageRelease.Enabled = true;
                if (addDevicesAgregate.signalContactArticle != null)
                    checkBoxSignalContact.Enabled = true;
                if (addDevicesAgregate.auxiliaryContactArticle != null)
                    checkBoxAuxiliaryContact.Enabled = true;
                if (addDevicesAgregate.signalOrAuxiliaryContactArticle != null)
                    checkBoxSignalOrAuxiliaryContact.Enabled = true;
            }
        }

        private void WriteDeviceData(string article, int rowOffset)
        {
            if (string.IsNullOrEmpty(article)) return;
            var writeExcel = new WriteExcel(dataInXml, addDevicesAgregate.vendor, rowOffset, article);
            writeExcel.Start();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            int maxRows = 1000;
            int rowsLine = currentRow;

            while (rowsLine < maxRows && Worksheet.Cells[rowsLine, 1].Value2 != null || Worksheet.Cells[rowsLine + 1, 1].Value2 != null)
            {
                ++rowsLine;
            }

            if (checkBoxShuntTrip24V.Checked) WriteDeviceData(addDevicesAgregate.shuntTrip24vArticle, rowsLine++ - currentRow);

            if (checkBoxShuntTrip48V.Checked) WriteDeviceData(addDevicesAgregate.shuntTrip48vArticle, rowsLine++ - currentRow);

            if (checkBoxShuntTrip230V.Checked) WriteDeviceData(addDevicesAgregate.shuntTrip230vArticle, rowsLine++ - currentRow);

            if (checkBoxUndervoltageRelease.Checked) WriteDeviceData(addDevicesAgregate.undervoltageReleaseArticle, rowsLine++ - currentRow);
           
            if (checkBoxSignalContact.Checked) WriteDeviceData(addDevicesAgregate.signalContactArticle, rowsLine++ - currentRow);

            if (checkBoxAuxiliaryContact.Checked) WriteDeviceData(addDevicesAgregate.auxiliaryContactArticle, rowsLine++ - currentRow);

            if (checkBoxSignalOrAuxiliaryContact.Checked) WriteDeviceData(addDevicesAgregate.signalOrAuxiliaryContactArticle, rowsLine++ - currentRow);

            Close();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
