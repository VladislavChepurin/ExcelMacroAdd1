using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Forms.ViewModels;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services.Interfaces;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class AdditionalDevicesForm : Form    {         
              
        private readonly AdditionalDevicesViewModel additionalDevicesViewModel;

        private void AdditionalDevicesForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            SelectionModularDevices main = this.Owner as SelectionModularDevices;
            main?.Show();
        }

        public AdditionalDevicesForm(IDataInXml dataInXml, IAdditionalModularDevicesData accessData, IFormSettings formSettings)
        {
            TopMost = formSettings.FormTopMost;
            additionalDevicesViewModel = new AdditionalDevicesViewModel(dataInXml, accessData);
            InitializeComponent();           
            InitializeDataBindings();
            
            additionalDevicesViewModel.RequestClose += (s, e) => this.Close();
            this.Load += (s, e) => additionalDevicesViewModel.Start();
            btnClose.Click += (s, e) => this.Close();           
            btnApply.Click += (s, e) => additionalDevicesViewModel.HandleBtnApplyClick();
        }

        private void InitializeDataBindings()
        {
            var checkBoxBindings = new Dictionary<System.Windows.Forms.CheckBox, string>
            {
                { checkBoxShuntTrip24V, "ShuntTrip24V" },
                { checkBoxShuntTrip48V, "ShuntTrip48V" },
                { checkBoxShuntTrip230V, "ShuntTrip230V" },
                { checkBoxUndervoltageRelease, "UndervoltageRelease" },
                { checkBoxSignalContact, "SignalContact" },
                { checkBoxAuxContact, "AuxContact" },
                { checkBoxSignalOrAuxContact, "SignalOrAuxContact" }
            };

            foreach (var checkBox in checkBoxBindings)
            {
                // Привязка для свойства Checked
                checkBox.Key.DataBindings.Add(
                    "Checked",
                    additionalDevicesViewModel,
                    $"IsActive{checkBox.Value}",
                    false,
                    DataSourceUpdateMode.OnPropertyChanged
                );

                // Привязка для свойства Enabled
                checkBox.Key.DataBindings.Add(
                    "Enabled",
                    additionalDevicesViewModel,
                    $"IsEnabled{checkBox.Value}",
                    false,
                    DataSourceUpdateMode.OnPropertyChanged
                );

                // Опционально: Добавляем валидацию
                checkBox.Key.DataBindings[0].FormattingEnabled = true;
                checkBox.Key.DataBindings[0].NullValue = false;
                checkBox.Key.DataBindings[1].FormattingEnabled = true;
            }

            btnApply.DataBindings.Add("Enabled",
                additionalDevicesViewModel,
                nameof(additionalDevicesViewModel.IsEnabledBtnApply),
                false,
                DataSourceUpdateMode.OnPropertyChanged
             );

            label1.DataBindings.Add("Visible",
                additionalDevicesViewModel,
                nameof(additionalDevicesViewModel.IsVisibleLabel),
                false,
                DataSourceUpdateMode.OnPropertyChanged      
             );      
         }
    }
}
