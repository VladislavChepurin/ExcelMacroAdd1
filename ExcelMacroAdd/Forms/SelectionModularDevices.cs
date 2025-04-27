using ExcelMacroAdd.BisinnesLayer;
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
        static readonly Mutex Mutex = new Mutex(false, "MutexSelectionModularDevices_SingleInstance");
        private bool _mutexAcquired = false;
           
        public SelectionModularDevices(IDataInXml dataInXml, AccessData accessData, IFormSettings formSettings)
        {
            InitializeComponent();
            try
            {
                _mutexAcquired = Mutex.WaitOne(TimeSpan.FromSeconds(1), false);
                if (!_mutexAcquired)
                {
                    Close();
                }
            }
            catch (AbandonedMutexException)
            {
                _mutexAcquired = true; // Мьютекс был оставлен, но теперь принадлежит текущему потоку
            }
            this.dataInXml = dataInXml;
            this.accessData = accessData;
            this.formSettings = formSettings;
            TopMost  = formSettings.FormTopMost;           
        }

        private void SelectionModularDevices_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (_mutexAcquired)
            {
                Mutex.ReleaseMutex();
                _mutexAcquired = false;
            }
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
