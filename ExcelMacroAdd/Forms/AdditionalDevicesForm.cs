using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Models;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Threading.Tasks;
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

        //Singelton
        private static AdditionalDevicesForm instance;
        public static void getInstance(IDataInXml dataInXml, IAccessAdditionalModularDevicesData accessData, IFormSettings formSettings)
        {
            if (instance == null)
            {
                instance = new AdditionalDevicesForm(dataInXml, accessData)
                {
                    TopMost = formSettings.FormTopMost
                };
                instance.ShowDialog();
            }
        }
        private void AdditionalDevicesForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            instance = null;
        }

        private AdditionalDevicesForm(IDataInXml dataInXml, IAccessAdditionalModularDevicesData accessData)
        {
            InitializeComponent();

            this.dataInXml = dataInXml;
            this.accessData = accessData;          

            currentRow = Cell.Row;
            string sArticle = Convert.ToString(Worksheet.Cells[currentRow, 1].Value2);

            if (sArticle == null)
                return;

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

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
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

            bool AllNull(params string[] strings)
            {
                return strings.All(s => s == null);
            }

            if (!additionalDevicesAgregateCheck)
            {
                button1.Enabled = true;
                label1.Visible = false;
                if (addDevicesAgregate.shuntTrip24vArticle != null)
                    checkBox1.Enabled = true;
                if (addDevicesAgregate.shuntTrip48vArticle != null)
                    checkBox2.Enabled = true;
                if (addDevicesAgregate.shuntTrip230vArticle != null)
                    checkBox3.Enabled = true;
                if (addDevicesAgregate.undervoltageReleaseArticle != null)
                    checkBox4.Enabled = true;
                if (addDevicesAgregate.signalContactArticle != null)
                    checkBox5.Enabled = true;
                if (addDevicesAgregate.auxiliaryContactArticle != null)
                    checkBox6.Enabled = true;
                if (addDevicesAgregate.signalOrAuxiliaryContactArticle != null)
                    checkBox7.Enabled = true;
            }           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int rowsLine = currentRow;

            while (Worksheet.Cells[rowsLine, 1].Value2 != null || Worksheet.Cells[rowsLine+1, 1].Value2 != null)
            {
                ++rowsLine;
            }

            if (checkBox1.Checked)
            {
                var writeExcel = new WriteExcel(dataInXml, addDevicesAgregate.vendor, rowsLine++ - currentRow, addDevicesAgregate.shuntTrip24vArticle);
                writeExcel.Start();
            }

            if (checkBox2.Checked)
            {
                var writeExcel = new WriteExcel(dataInXml, addDevicesAgregate.vendor, rowsLine++ - currentRow, addDevicesAgregate.shuntTrip48vArticle);
                writeExcel.Start();
            }

            if (checkBox3.Checked)
            {
                var writeExcel = new WriteExcel(dataInXml, addDevicesAgregate.vendor, rowsLine++ - currentRow, addDevicesAgregate.shuntTrip230vArticle);
                writeExcel.Start();
            }

            if (checkBox4.Checked)
            {
                var writeExcel = new WriteExcel(dataInXml, addDevicesAgregate.vendor, rowsLine++ - currentRow, addDevicesAgregate.undervoltageReleaseArticle);
                writeExcel.Start();
            }

            if (checkBox5.Checked)
            {
                var writeExcel = new WriteExcel(dataInXml, addDevicesAgregate.vendor, rowsLine++ - currentRow, addDevicesAgregate.signalContactArticle);
                writeExcel.Start();
            }

            if (checkBox6.Checked)
            {
                var writeExcel = new WriteExcel(dataInXml, addDevicesAgregate.vendor, rowsLine++ - currentRow, addDevicesAgregate.auxiliaryContactArticle);
                writeExcel.Start();
            }

            if (checkBox7.Checked)
            {
                var writeExcel = new WriteExcel(dataInXml, addDevicesAgregate.vendor, rowsLine++ - currentRow, addDevicesAgregate.signalOrAuxiliaryContactArticle);
                writeExcel.Start();
            }
            Close();
        }
    }
}
