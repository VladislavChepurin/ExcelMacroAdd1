using ExcelMacroAdd.BisinnesLayer;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services.Interfaces;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class SelectionModularDevices : Form
    {
        private readonly IDataInXml dataInXml;
        private readonly AccessData accessData;
        private readonly IFormSettings formSettings;

        private SelectionModularDevices(IDataInXml dataInXml, AccessData accessData, IFormSettings formSettings)
        {
            this.dataInXml = dataInXml;
            this.accessData = accessData;
            this.formSettings = formSettings;
            InitializeComponent();
        }

        //Singelton
        private static SelectionModularDevices instance;
        public static async Task getInstance(IDataInXml dataInXml, AccessData accessData, IFormSettings formSettings)
        {
            if (instance == null)
            {
                await Task.Run(() =>
                {
                    instance = new SelectionModularDevices(dataInXml, accessData, formSettings)
                    {
                        TopMost = formSettings.FormTopMost
                    };
                    instance.ShowDialog();
                });
            }
        }
        private void SelectionModularDevices_FormClosed(object sender, FormClosedEventArgs e) =>
              instance = null;

        private void SelectionModularDevices_Load(object sender, EventArgs e)
        {
            button1.Click += async (s, a) =>
            {
                await SelectionCircuitBreaker.getInstance(dataInXml, accessData, formSettings);                
            };

            button2.Click += async (s, a) =>
            {
                await SelectionSwitch.getInstance(dataInXml, accessData, formSettings);
            };

            button3.Click += async (s, a) =>
            {
                await AdditionalDevicesForm.getInstance(dataInXml, accessData, formSettings);
            };
        }
    }
}
