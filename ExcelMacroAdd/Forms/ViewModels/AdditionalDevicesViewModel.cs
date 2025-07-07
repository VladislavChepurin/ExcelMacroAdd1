using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Models;
using ExcelMacroAdd.Services;
using ExcelMacroAdd.Services.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelMacroAdd.Forms.ViewModels
{
    public class AdditionalDevicesViewModel: AbstractFunctions, INotifyPropertyChanged
    {
        private readonly IDataInXml _dataInXml;
        private readonly IAdditionalModularDevicesData _accessData;
        private readonly int _currentRow;
        private readonly string _article;
        private readonly AdditionalDevices _circuitBreakerData;
        private readonly AdditionalDevices _switchData;
        private readonly AdditionalDevices _addDevicesAgregate;       

        #region Property isEnabled CheckBox

        private bool _isEnabledShuntTrip24V;
        private bool _isEnabledShuntTrip48V;
        private bool _isEnabledShuntTrip230V;
        private bool _isEnabledUndervoltageRelease;
        private bool _isEnabledSignalContact;
        private bool _isEnabledAuxContact;
        private bool _isEnabledSignalOrAuxContact;

        public bool IsEnabledShuntTrip24V
        {
            get => _isEnabledShuntTrip24V;
            set { _isEnabledShuntTrip24V = value; OnPropertyChanged(nameof(IsEnabledShuntTrip24V)); }
        }

        public bool IsEnabledShuntTrip48V
        {
            get => _isEnabledShuntTrip48V;
            set { _isEnabledShuntTrip48V = value; OnPropertyChanged(nameof(IsEnabledShuntTrip48V)); }
        }

        public bool IsEnabledShuntTrip230V
        {
            get => _isEnabledShuntTrip230V;
            set { _isEnabledShuntTrip230V = value; OnPropertyChanged(nameof(IsEnabledShuntTrip230V)); }
        }

        public bool IsEnabledUndervoltageRelease
        {
            get => _isEnabledUndervoltageRelease;
            set { _isEnabledUndervoltageRelease = value; OnPropertyChanged(nameof(IsEnabledUndervoltageRelease)); }
        }

        public bool IsEnabledSignalContact
        {
            get => _isEnabledSignalContact;
            set { _isEnabledSignalContact = value; OnPropertyChanged(nameof(IsEnabledSignalContact)); }
        }

        public bool IsEnabledAuxContact
        {
            get => _isEnabledAuxContact;
            set { _isEnabledAuxContact = value; OnPropertyChanged(nameof(IsEnabledAuxContact)); }
        }

        public bool IsEnabledSignalOrAuxContact
        {
            get => _isEnabledSignalOrAuxContact;
            set { _isEnabledSignalOrAuxContact = value; OnPropertyChanged(nameof(IsEnabledSignalOrAuxContact)); }
        }

        #endregion

        #region Property isActive CheckBox

        private bool _isActiveShuntTrip24V;
        private bool _isActiveShuntTrip48V;
        private bool _isActiveShuntTrip230V;
        private bool _isActiveUndervoltageRelease;
        private bool _isActiveSignalContact;
        private bool _isActiveAuxContact;
        private bool _isActiveSignalOrAuxContact;

        public bool IsActiveShuntTrip24V
        {
            get => _isActiveShuntTrip24V;
            set { _isActiveShuntTrip24V = value; OnPropertyChanged(nameof(IsActiveShuntTrip24V)); }
        }

        public bool IsActiveShuntTrip48V
        {
            get => _isActiveShuntTrip48V;
            set { _isActiveShuntTrip48V = value; OnPropertyChanged(nameof(IsActiveShuntTrip48V)); }
        }

        public bool IsActiveShuntTrip230V
        {
            get => _isActiveShuntTrip230V;
            set { _isActiveShuntTrip230V = value; OnPropertyChanged(nameof(IsActiveShuntTrip230V)); }
        }

        public bool IsActiveUndervoltageRelease
        {
            get => _isActiveUndervoltageRelease;
            set { _isActiveUndervoltageRelease = value; OnPropertyChanged(nameof(IsActiveUndervoltageRelease)); }
        }

        public bool IsActiveSignalContact
        {
            get => _isActiveSignalContact;
            set { _isActiveSignalContact = value; OnPropertyChanged(nameof(IsActiveSignalContact)); }
        }

        public bool IsActiveAuxContact
        {
            get => _isActiveAuxContact;
            set { _isActiveAuxContact = value; OnPropertyChanged(nameof(IsActiveAuxContact)); }
        }

        public bool IsActiveSignalOrAuxContact
        {
            get => _isActiveSignalOrAuxContact;
            set { _isActiveSignalOrAuxContact = value; OnPropertyChanged(nameof(IsActiveSignalOrAuxContact)); }
        }

        #endregion

        #region Property ButtonApply and Label

        private bool _isEnabledBtnApply;
        private bool _isVisibleLabel;

        public bool IsEnabledBtnApply
        {
            get => _isEnabledBtnApply;
            set { _isEnabledBtnApply = value; OnPropertyChanged(nameof(IsEnabledBtnApply)); }
        }

        public bool IsVisibleLabel
        {
            get => _isVisibleLabel;
            set { _isVisibleLabel = value; OnPropertyChanged(nameof(IsVisibleLabel)); }
        }

        #endregion

        private bool AllNull(params string[] values)
        {
            if (values == null) return true;
            return values.All(v => string.IsNullOrEmpty(v));
        }

        public AdditionalDevicesViewModel(IDataInXml dataInXml, IAdditionalModularDevicesData accessData)
        {
            this._dataInXml = dataInXml;
            this._accessData = accessData;       

            _currentRow = Cell.Row;

            _article = Convert.ToString(Worksheet.Cells[_currentRow, 1].Value2);
            if (_article != null)
            {
                _circuitBreakerData = accessData.AccessAdditionalModularDevices.GetEntityAdditionalCircuitBreaker(_article);
                _switchData = accessData.AccessAdditionalModularDevices.GetEntityAdditionalSwitch(_article);
                _addDevicesAgregate = (_circuitBreakerData.vendor != null) ? _circuitBreakerData : _switchData;
            }
        }
             
        public override void Start()
        {
            IsVisibleLabel = true;

            if (_addDevicesAgregate == null)
                return;

            bool additionalDevicesAgregateCheck = AllNull(
                _addDevicesAgregate.shuntTrip24vArticle,
                _addDevicesAgregate.shuntTrip48vArticle,
                _addDevicesAgregate.shuntTrip230vArticle,
                _addDevicesAgregate.undervoltageReleaseArticle,
                _addDevicesAgregate.signalContactArticle,
                _addDevicesAgregate.auxiliaryContactArticle,
                _addDevicesAgregate.signalOrAuxiliaryContactArticle);

            if (!additionalDevicesAgregateCheck)
            {
                IsEnabledBtnApply = true;
                IsVisibleLabel = false;

                IsEnabledShuntTrip24V = !string.IsNullOrEmpty(_addDevicesAgregate.shuntTrip24vArticle);
                IsEnabledShuntTrip48V = !string.IsNullOrEmpty(_addDevicesAgregate.shuntTrip48vArticle);
                IsEnabledShuntTrip230V = !string.IsNullOrEmpty(_addDevicesAgregate.shuntTrip230vArticle);
                IsEnabledUndervoltageRelease = !string.IsNullOrEmpty(_addDevicesAgregate.undervoltageReleaseArticle);
                IsEnabledSignalContact = !string.IsNullOrEmpty(_addDevicesAgregate.signalContactArticle);
                IsEnabledAuxContact = !string.IsNullOrEmpty(_addDevicesAgregate.auxiliaryContactArticle);
                IsEnabledSignalOrAuxContact = !string.IsNullOrEmpty(_addDevicesAgregate.signalOrAuxiliaryContactArticle);
            }
        }
      
        public void HandleBtnApplyClick()
        {
            // Создаем список условий и соответствующих артикулов
            var deviceActions = new List<(Func<bool> Condition, string Article)>
            {
                (() => IsActiveShuntTrip24V, _addDevicesAgregate.shuntTrip24vArticle),
                (() => IsActiveShuntTrip48V, _addDevicesAgregate.shuntTrip48vArticle),
                (() => IsActiveShuntTrip230V, _addDevicesAgregate.shuntTrip230vArticle),
                (() => IsActiveUndervoltageRelease, _addDevicesAgregate.undervoltageReleaseArticle),
                (() => IsActiveSignalContact, _addDevicesAgregate.signalContactArticle),
                (() => IsActiveAuxContact, _addDevicesAgregate.auxiliaryContactArticle),
                (() => IsActiveSignalOrAuxContact, _addDevicesAgregate.signalOrAuxiliaryContactArticle)
            };

            Task.Run(() =>
            {
                
                int rowOffset = 0;
                foreach (var action in deviceActions)
                {
                    if (action.Condition())
                    {
                        WriteDeviceData(action.Article, ++rowOffset);
                    }
                }
            })
                .ContinueWith(t =>
                {
                    // Этот код выполнится после завершения фоновой задачи
                    if (t.IsCompleted)
                    {                      
                        RequestClose?.Invoke(this, EventArgs.Empty);
                    }
                    else if (t.IsFaulted)
                    {
                        // Обработка ошибок
                        Logger.LogException(t.Exception);
                    }
                }, TaskScheduler.FromCurrentSynchronizationContext()); // Гарантирует выполнение в UI потоке                     
        }

        private void WriteDeviceData(string article, int rowOffset)
        {
            if (string.IsNullOrEmpty(article)) return;
            var writeExcel = new WriteExcel(_dataInXml, _addDevicesAgregate.vendor, article, rowOffset);
            writeExcel.Start();
        }


        public event EventHandler RequestClose;
              
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }       
    }
}
