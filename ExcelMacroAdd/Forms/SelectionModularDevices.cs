using ExcelMacroAdd.BusinessLayer;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services.Interfaces;
using System;
using System.Threading;
using System.Windows.Forms;

//Rewiew OK 21.04.2025
namespace ExcelMacroAdd.Forms
{
    public partial class SelectionModularDevices : Form
    {
        private readonly IDataInXml dataInXml;
        private readonly AccessData accessData;
        private readonly IFormSettings formSettings;
           
        public SelectionModularDevices(IDataInXml dataInXml, AccessData accessData, IFormSettings formSettings)
        {
            InitializeComponent();           
            this.dataInXml = dataInXml;
            this.accessData = accessData;
            this.formSettings = formSettings;
            TopMost  = formSettings.FormTopMost;           
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
                using (var form = new SelectionCircuitBreaker(dataInXml, accessData, formSettings))                
                    ShowChildForm(form);
            };

            btnSelectionSwitchShow.Click += (s, a) =>
            {
                using (var form = new SelectionSwitch(dataInXml, accessData, formSettings))
                    ShowChildForm(form);
            };

            btnAdditionalDevicesShow.Click += (s, a) =>
            {
                using (var form = new AdditionalDevicesForm(dataInXml, accessData, formSettings))
                    ShowChildForm(form);
            };
        }
    }
}
