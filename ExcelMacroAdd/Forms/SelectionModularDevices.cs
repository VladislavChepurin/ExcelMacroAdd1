using ExcelMacroAdd.BisinnesLayer;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services.Interfaces;
using System;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class SelectionModularDevices : Form
    {
        private readonly IDataInXml dataInXml;
        private readonly AccessData accessData;
        private readonly IFormSettings formSettings;
        static readonly Mutex Mutex = new Mutex(false, "MutexSelectionModularDevices_SingleInstance");

        public SelectionModularDevices(IDataInXml dataInXml, AccessData accessData, IFormSettings formSettings)
        {
            // проверяем, находится ли мьютекс в сигнальном состоянии
            var signaled = Mutex.WaitOne(TimeSpan.FromSeconds(1), false);

            // Если состояние несигнальное (несвободное) — значит другой поток уже владеет мьютексом
            if (!signaled)
            {
                Close();
            }

            this.dataInXml = dataInXml;
            this.accessData = accessData;
            this.formSettings = formSettings;
            TopMost  = formSettings.FormTopMost;
            InitializeComponent();
        }

        private void SelectionModularDevices_FormClosed(object sender, FormClosedEventArgs e) =>
            Mutex.ReleaseMutex();

        private void SelectionModularDevices_Load(object sender, EventArgs e)
        {
            button1.Click += (s, a) =>
            {
                Hide();
                var selectionCircuitBreaker = new SelectionCircuitBreaker(dataInXml, accessData, formSettings)
                {
                    Owner = this
                };
                selectionCircuitBreaker.ShowDialog();
            };

            button2.Click += (s, a) =>
            {
                Hide();
                var selectionSwitch = new SelectionSwitch(dataInXml, accessData, formSettings)
                {
                    Owner = this
                };
                selectionSwitch.ShowDialog();
            };

            button3.Click += (s, a) =>
            {
                Hide();
                var additionalDevicesForm = new AdditionalDevicesForm(dataInXml, accessData, formSettings)
                {
                    Owner = this
                };
                additionalDevicesForm.ShowDialog();
            };
        }
    }
}
