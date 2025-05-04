using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Forms.ViewModels;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using static ExcelMacroAdd.Forms.ViewModels.TermoCalculationViewModel;

namespace ExcelMacroAdd.Forms
{
    public partial class TermoCalculation : Form
    {
        private readonly TermoCalculationViewModel termoCalculationViewModel;
        static readonly Mutex Mutex = new Mutex(false, "MutexTermoCalculation_SingleInstance");
        private bool _mutexAcquired = false;
               
        public TermoCalculation(ITermoCalcData accessData, IFormSettings formSettings)
        {
            termoCalculationViewModel = new TermoCalculationViewModel(accessData);

            InitializeComponent();
            InitializeDataBindings();
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

            TopMost = formSettings.FormTopMost;
            this.Load += (s, e) => termoCalculationViewModel.Start();            
            textBoxHeight.KeyPress += NumericTextBox_KeyPress;
            textBoxWidth.KeyPress += NumericTextBox_KeyPress;
            textBoxDepth.KeyPress += NumericTextBox_KeyPress;
            textBoxHeatDissipation.KeyPress += NumericTextBox_KeyPress;
            textBoxMinTemp.KeyPress += NumericTextBox_KeyPressSigned;
            textBoxTargetTemp.KeyPress += NumericTextBox_KeyPressSigned;        
            btnAplly.Click += (s, e) => termoCalculationViewModel.BthApllyClick();
        }

        private void NumericTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            var textBox = sender as TextBox;

            // Разрешаем:
            // - Цифры
            // - Backspace
            // - Delete           
            bool isDigit = char.IsDigit(e.KeyChar);
            bool isControl = char.IsControl(e.KeyChar);  
            
            e.Handled = !(isDigit || isControl);
        }

        private void NumericTextBox_KeyPressSigned(object sender, KeyPressEventArgs e)
        {
            var textBox = sender as TextBox;

            // Разрешаем:
            // - Цифры
            // - Знак минуса только в начале
            // - Backspace
            // - Delete
            bool isDigit = char.IsDigit(e.KeyChar);
            bool isControl = char.IsControl(e.KeyChar);
            bool isMinus = e.KeyChar == '-';
            bool canInsertMinus = textBox.SelectionStart == 0 && !textBox.Text.Contains('-');

            e.Handled = !(isDigit || isControl || (isMinus && canInsertMinus));
        }

        private void InitializeDataBindings()
        {
            var textBoxBindings = new Dictionary<System.Windows.Forms.TextBox, string>
            {
                { textBoxHeight, "Height" },
                { textBoxWidth, "Width" },
                { textBoxDepth, "Depth" },
                { textBoxMinTemp, "MinTemp" },
                { textBoxTargetTemp, "TargetTemp" },
                { textBoxDifferenceTemp, "DifferenceTemp" },
                { textBoxHeatDissipation, "HeatDissipation" }
            };

            foreach (var textBox in textBoxBindings)
            {
                // Привязка для свойства Text
                textBox.Key.DataBindings.Add(
                    "Text",
                    termoCalculationViewModel,
                    $"TextBox{textBox.Value}",
                    false,
                    DataSourceUpdateMode.OnPropertyChanged
                );
            }                     

            AddRadioButtonBinding(radioComprehensiveAccess, Options.ComprehensiveAccess);
            AddRadioButtonBinding(radioPlacementWall, Options.PlacementWall);
            AddRadioButtonBinding(radioLastRow, Options.LastRow);
            AddRadioButtonBinding(radioLastRowNearWall, Options.LastRowNearWall);
            AddRadioButtonBinding(radioMiddleRow, Options.MiddleRow);
            AddRadioButtonBinding(radioMiddleRowNearWall, Options.MiddleRowNearWall);
            AddRadioButtonBinding(radioMiddleRowNearWallWithUpperClosed, Options.MiddleRowNearWallWithUpperClosed);

            labelResult.DataBindings.Add
            (
                "Text",
                termoCalculationViewModel,
                nameof(termoCalculationViewModel.TextResultLabel),
                false,
                DataSourceUpdateMode.OnPropertyChanged
            );

            labelResult.DataBindings.Add
            (
                "Visible",
                termoCalculationViewModel,
                nameof(termoCalculationViewModel.IsVisibleLabel),         
                false,
                DataSourceUpdateMode.OnPropertyChanged
            );

            comboBoxInstallation.DataBindings.Add
            (
                "SelectedIndex",
                termoCalculationViewModel,
                nameof(termoCalculationViewModel.InstallationIndex),
                false,
                DataSourceUpdateMode.OnPropertyChanged
            );


            comboBoxCabinetMaterial.DataBindings.Add
            (
                "SelectedIndex",
                termoCalculationViewModel,
                nameof(termoCalculationViewModel.CabinetMaterialIndex),
                false,
                DataSourceUpdateMode.OnPropertyChanged
            );

            comboBoxInsulationМaterial.DataBindings.Add
            (
                "SelectedIndex",
                termoCalculationViewModel,
                nameof(termoCalculationViewModel.InsulationМaterialIndex),
                false,
                DataSourceUpdateMode.OnPropertyChanged
            );
        }

        private void AddRadioButtonBinding(RadioButton radioButton, Options optionValue)
        {
            var binding = new Binding(
                "Checked",
                termoCalculationViewModel,
                nameof(termoCalculationViewModel.SelectedOption),
                false, // форматирование включено
                DataSourceUpdateMode.OnPropertyChanged);

            binding.Format += (sender, e) =>
            {
                e.Value = e.Value != null && (Options)e.Value == optionValue;
            };

            binding.Parse += (sender, e) =>
            {
                if ((bool)e.Value)
                {
                    termoCalculationViewModel.SelectedOption = optionValue;
                }
            };

            radioButton.DataBindings.Add(binding);
        }

        private void TermoCalculation_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (_mutexAcquired)
            {
                Mutex.ReleaseMutex();
                _mutexAcquired = false;
            }
        }                              
    }
}
