using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services.Interfaces;
using System;
using System.Windows.Forms;

//Rewiew OK 21.04.2025
namespace ExcelMacroAdd.Forms
{
    public partial class SelectionModularDevices : Form
    {
        private readonly IDataInXml dataInXml;
        private readonly ICircuitBreakerService circuitBreakerService;
        private readonly ISwitchService switchService;
        private readonly IAdditionalDevicesService additionalDevicesService;
        private readonly IFormSettings formSettings;

        public SelectionModularDevices(
            IDataInXml dataInXml,
            ICircuitBreakerService circuitBreakerService,
            ISwitchService switchService,
            IAdditionalDevicesService additionalDevicesService,
            IFormSettings formSettings)
        {
            InitializeComponent();
            this.dataInXml = dataInXml;
            this.circuitBreakerService = circuitBreakerService;
            this.switchService = switchService;
            this.additionalDevicesService = additionalDevicesService;
            this.formSettings = formSettings;
            TopMost = formSettings.FormTopMost;
        }

        private void ShowChildForm(Form childForm)
        {
            childForm.FormClosed += (s, e) => Show();
            Hide();
            childForm.ShowDialog();
        }

        private void SelectionModularDevices_Load(object sender, EventArgs e)
        {
            btnSelectionCircuitBreakerShow.Click += (s, a) =>
            {
                using (var form = new SelectionCircuitBreaker(dataInXml, circuitBreakerService, formSettings))
                {
                    ShowChildForm(form);
                }
            };

            btnSelectionSwitchShow.Click += (s, a) =>
            {
                using (var form = new SelectionSwitch(dataInXml, switchService, formSettings))
                {
                    ShowChildForm(form);
                }
            };

            btnAdditionalDevicesShow.Click += (s, a) =>
            {
                using (var form = new AdditionalDevicesForm(dataInXml, additionalDevicesService, formSettings))
                {
                    ShowChildForm(form);
                }
            };
        }
    }
}
