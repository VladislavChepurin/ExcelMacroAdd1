using ExcelMacroAdd.Forms.ViewModels;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    internal partial class FillingOutPassports : Form
    {
        private readonly FillingOutPassportViewModel fillingOutPassportViewModel;

        public FillingOutPassports(IFillingOutThePassportSettings resources)
        {
            fillingOutPassportViewModel = new FillingOutPassportViewModel(resources);
            InitializeComponent();
            InitializeDataBindings();

            fillingOutPassportViewModel.RequestClose += (s, e) => this.Close();
            this.Load += (s, e) => fillingOutPassportViewModel.Start();
            btnClose.Click += (s, e) => this.Close();
        }

        private void InitializeDataBindings()
        {
            infoLabel.DataBindings.Add("Text",
               fillingOutPassportViewModel,
               nameof(fillingOutPassportViewModel.InfoLabelText),
               false,
               DataSourceUpdateMode.OnPropertyChanged
            );

            btnClose.DataBindings.Add("Enabled",
               fillingOutPassportViewModel,
               nameof(fillingOutPassportViewModel.IsEnabledBtnClose),             
               false,
               DataSourceUpdateMode.OnPropertyChanged
            );

            progressBar.DataBindings.Add("Minimum",
               fillingOutPassportViewModel,
               nameof(fillingOutPassportViewModel.ProgressBarMinimum),
               false,
               DataSourceUpdateMode.OnPropertyChanged
            );

            progressBar.DataBindings.Add("Maximum",
               fillingOutPassportViewModel,
               nameof(fillingOutPassportViewModel.ProgressBarMaximum),
               false,
               DataSourceUpdateMode.OnPropertyChanged
            );

            progressBar.DataBindings.Add("Step",
              fillingOutPassportViewModel,
              nameof(fillingOutPassportViewModel.ProgressBarStep),
              false,
              DataSourceUpdateMode.OnPropertyChanged
           );

            progressBar.DataBindings.Add("Value",
              fillingOutPassportViewModel,
              nameof(fillingOutPassportViewModel.ProgressBarValue),
              false,
              DataSourceUpdateMode.OnPropertyChanged
           );                        
        }
    }
}
