using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Services;
using ExcelMacroAdd.BusinessLayer.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms.ViewModels
{
    public class NotPriceComponentsViewModel : AbstractFunctions, INotifyPropertyChanged
    {
        private const int ArticleColumn = 1;
        private const int IsDiscontinued = 2;
        private const int DescriptionColumn = 2;
        private const int QuantityColumn = 3;
        private const int MultiplicityColumn = 4;
        private const int ProductVendorColumn = 5;
        private const int DiscountColumn = 6;
        private const int PriceColumn = 7;
        private const int TotalPriceColumn = 8;
        private const int CoastColumn = 9;
        private const int DateColumn = 10;
        private const int LinkColumn = 11;
        private const int MaxDisplayItems = 1000;
        private const int FilterDelayMs = 300;

        private readonly INotPriceComponentsService _notPriceComponentsService;
        private BindingList<NotPriceComponent> _filteredList;
        private List<NotPriceComponent> _allRecords;
        private NotPriceComponent _selectedRecord;
        private string _searchTerm;
        private CancellationTokenSource _filterTokenSource;
        private bool _isLoading;
        private string _countStatusList;
        private string _linkToTheWebsite = string.Empty;

        private SynchronizationContext _uiContext;

        public string CountStatusList
        {
            get => _countStatusList;
            set
            {
                _countStatusList = value;
                OnPropertyChanged(nameof(CountStatusList));
            }
        }

        public string LinkToTheWebsite
        {
            get => _linkToTheWebsite;
            set
            {
                _linkToTheWebsite = value;
                OnPropertyChanged(nameof(LinkToTheWebsite));
                OnPropertyChanged(nameof(DisplayLink));
            }
        }

        public string DisplayLink
        {
            get
            {
                if (string.IsNullOrWhiteSpace(LinkToTheWebsite))
                    return String.Empty;

                if (LinkToTheWebsite.Length > 60)
                {
                    return LinkToTheWebsite.Substring(0, 57) + "...";
                }

                return LinkToTheWebsite;
            }
        }

        public BindingList<NotPriceComponent> FilteredList
        {
            get => _filteredList;
            set
            {
                if (_filteredList != value)
                {
                    _filteredList = value;
                    OnPropertyChanged(nameof(FilteredList));
                }
            }
        }


        public NotPriceComponent SelectedRecord
        {
            get => _selectedRecord;
            set
            {
                if (_selectedRecord != value)
                {
                    _selectedRecord = value;
                    OnPropertyChanged(nameof(SelectedRecord));
                    LinkToTheWebsite = _selectedRecord?.Link ?? string.Empty;
                }
            }
        }

        public string SearchTerm
        {
            get => _searchTerm;
            set
            {
                if (_searchTerm != value)
                {
                    _searchTerm = value;
                    OnPropertyChanged(nameof(SearchTerm));
                    ApplyFilterAsync();
                }
            }
        }

        public bool IsLoading
        {
            get => _isLoading;
            set
            {
                if (_isLoading != value)
                {
                    _isLoading = value;
                    OnPropertyChanged(nameof(IsLoading));
                }
            }
        }

        public void OpenLink()
        {
            if (!string.IsNullOrWhiteSpace(LinkToTheWebsite))
            {
                try
                {
                    string url = LinkToTheWebsite;
                    if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                    {
                        url = "http://" + url;
                    }

                    Process.Start(new ProcessStartInfo
                    {
                        FileName = url,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    // Обработка ошибок
                    CountStatusList = $"Ошибка открытия ссылки: {ex.Message}";
                }
            }
        }

        public NotPriceComponentsViewModel(INotPriceComponentsService notPriceComponentsService)
        {
            _notPriceComponentsService = notPriceComponentsService ?? throw new ArgumentNullException(nameof(notPriceComponentsService));
            _filterTokenSource = new CancellationTokenSource();
            _allRecords = new List<NotPriceComponent>();
            FilteredList = new BindingList<NotPriceComponent>();
            _uiContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();
        }

        public override async Task StartAsync()
        {
            try
            {
                IsLoading = true;
                var records = await _notPriceComponentsService.GetAllRecordsAsync().ConfigureAwait(false);
                _allRecords = records.ToList();
                FilteredList = new BindingList<NotPriceComponent>(_allRecords);
                CountStatusList = $"Всего доступно {_allRecords.Count} записей, выбрано {FilteredList.Count} записей";
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка загрузки данных: {ex.Message}", "Ошибка загрузки");
                Logger.LogException(ex);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async void ApplyFilterAsync()
        {
            _filterTokenSource.Cancel();
            _filterTokenSource = new CancellationTokenSource();
            var token = _filterTokenSource.Token;

            try
            {
                await Task.Delay(FilterDelayMs, token).ConfigureAwait(false);
                if (token.IsCancellationRequested) return;

                var search = SearchTerm?.Trim();
                List<NotPriceComponent> result;

                if (string.IsNullOrWhiteSpace(search))
                {
                    result = _allRecords;
                }
                else
                {
                    result = _allRecords
                    .AsParallel()
                    .WithCancellation(token)
                    .Where(item =>
                    item != null &&
                    (
                        (!string.IsNullOrEmpty(item.Article) &&
                         item.Article.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        (!string.IsNullOrEmpty(item.Description) &&
                         item.Description.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        (!string.IsNullOrEmpty(item.VendorDisplayName) &&
                         item.VendorDisplayName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
                    ))
                    .Take(MaxDisplayItems)
                    .ToList();
                }

                FilteredList = new BindingList<NotPriceComponent>(result);

                CountStatusList = $"Всего доступно {_allRecords.Count} записей, выбрано {FilteredList.Count} записей";
            }
            catch (TaskCanceledException)
            {
                // Фильтрация была отменена
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка фильтрации: {ex.Message}");
                Logger.LogException(ex);
            }
        }

        public void BtnWritingToSheet()
        {
            if (SelectedRecord == null)
            {
                MessageWarning("Пожалуйста, выберите запись для переноса в лист", "Запись не выбрана");
                return;
            }

            Range activeCell = null;
            try
            {
                activeCell = Worksheet.Application.ActiveCell;
                int currentRow = activeCell.Row;

                var selectedRecord = SelectedRecord;
                WriteToSheet(currentRow, selectedRecord);
                ActivateNextRow(currentRow);
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка при записи в лист: {ex.Message}", "Ошибка записи");
                Logger.LogException(ex);
            }
            finally
            {
                ReleaseComObjects(activeCell);
            }
        }

        private void WriteToSheet(int currentRow, NotPriceComponent record)
        {
            Worksheet.Cells[currentRow, ArticleColumn] = record.Article;
            Worksheet.Cells[currentRow, DescriptionColumn] = record.Description;
            Worksheet.Cells[currentRow, MultiplicityColumn] = record.MultiplicityDisplayName;
            Worksheet.Cells[currentRow, ProductVendorColumn] = record.VendorDisplayName;

            Worksheet.Cells[currentRow, DiscountColumn] = record.Discount;
            Worksheet.Cells[currentRow, DiscountColumn].NumberFormat = "0";

            // Записываем и форматируем цену
            Range priceCell = Worksheet.Cells[currentRow, PriceColumn];
            priceCell.Value2 = record.Price;
            priceCell.NumberFormat = "#,##0.00";

            Range totalPriceCell = Worksheet.Cells[currentRow, TotalPriceColumn];
            totalPriceCell.Formula = $"=G{currentRow}*(100-F{currentRow})/100";
            totalPriceCell.NumberFormat = "#,##0.00";

            Range coastCell = Worksheet.Cells[currentRow, CoastColumn];
            coastCell.Formula = $"=H{currentRow}*C{currentRow}";
            coastCell.NumberFormat = "#,##0.00";

            Worksheet.Cells[currentRow, DateColumn].NumberFormat = "ДД.ММ.ГГ ч:мм";
            Worksheet.Cells[currentRow, DateColumn] = DateTime.Now;
        }

        private void ActivateNextRow(int currentRow)
        {
            int nextRow = ++currentRow;
            Worksheet.Cells[nextRow, 1].Select();
        }

        private void SetCellValueWithFormat(Range cell, object value, string format)
        {
            cell.Value2 = value;
            cell.NumberFormat = format;
        }

        private void ReleaseComObjects(params object[] comObjects)
        {
            foreach (var obj in comObjects)
            {
                if (obj != null && Marshal.IsComObject(obj))
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
        }

        public async Task BtnAddRecord()
        {
            Range activeCell = null;
            try
            {
                activeCell = Worksheet.Application.ActiveCell;
                int currentRow = activeCell.Row;

                string article = GetCellValueAsString(Worksheet.Cells[currentRow, ArticleColumn]);

                if (string.IsNullOrWhiteSpace(article))
                {
                    MessageError("Добавить невозможно, пустой артикул", "Ошибка добавления");
                    return;
                }

                if (await _notPriceComponentsService.RecordExistsAsync(article).ConfigureAwait(false))
                {
                    MessageError($"Артикул {article} уже есть в базе данных", "Ошибка добавления");
                    return;
                }

                int discount = GetCellValueAsInt(Worksheet.Cells[currentRow, DiscountColumn]);
                string description = GetCellValueAsString(Worksheet.Cells[currentRow, DescriptionColumn]);
                string productVendorName = GetCellValueAsString(Worksheet.Cells[currentRow, ProductVendorColumn]);
                string multiplicityName = GetCellValueAsString(Worksheet.Cells[currentRow, MultiplicityColumn]);
                decimal price = GetCellValueAsDecimal(Worksheet.Cells[currentRow, PriceColumn]);
                string link = GetCellValueAsString(Worksheet.Cells[currentRow, LinkColumn]);

                if (string.IsNullOrEmpty(article) || string.IsNullOrEmpty(description) || string.IsNullOrEmpty(productVendorName))
                {
                    MessageWarning("Обязательные поля не заполнены", "Ошибка записи");
                    return;
                }

                await ProcessAddRecord(CreateSaveRequest(article, description, productVendorName, multiplicityName, price, discount, link));
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка при добавлении записи: {ex.Message}", "Ошибка добавления");
                Logger.LogException(ex);
            }
            finally
            {
                ReleaseComObjects(activeCell);
            }
        }

        private async Task ProcessAddRecord(NotPriceComponentSaveRequest request)
        {
            bool createVendorIfMissing = false;
            if (!await _notPriceComponentsService.VendorExistsAsync(request.ProductVendorName).ConfigureAwait(false))
            {
                if (!ConfirmAddNewVendor(request.ProductVendorName)) return;
                createVendorIfMissing = true;
            }

            await _notPriceComponentsService.AddRecordAsync(request, createVendorIfMissing).ConfigureAwait(false);
            await StartAsync();

            MessageInformation($"Успешно записано в базу данных!\nАртикул: {request.Article}\nВендор: {request.ProductVendorName}",
                "Запись успешна!");
        }

        private bool ConfirmAddNewVendor(string vendorName)
        {
            return MessageBox.Show(
                $"В БД вендора '{vendorName}' нет. Добавить нового вендора?",
                "Добавление нового вендора",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.Yes;
        }

        // Общий метод для обновления записи
        private void UpdateRecordInLists(NotPriceComponent updatedRecord)
        {
            _uiContext.Post(_ =>
            {
                bool wasSelected = SelectedRecord?.Id == updatedRecord.Id;

                // Обновляем в _allRecords
                var recordIndex = _allRecords.FindIndex(r => r.Id == updatedRecord.Id);
                if (recordIndex >= 0)
                {
                    _allRecords[recordIndex] = updatedRecord;
                }

                // Обновляем в FilteredList
                var filteredItem = FilteredList.FirstOrDefault(r => r.Id == updatedRecord.Id);
                if (filteredItem != null)
                {
                    var index = FilteredList.IndexOf(filteredItem);
                    FilteredList[index] = updatedRecord;
                }

                if (wasSelected)
                {
                    SelectedRecord = updatedRecord;
                }

                CountStatusList = $"Всего доступно {_allRecords.Count} записей, выбрано {FilteredList.Count} записей";
            }, null);
        }

        // Общий метод для удаления записи
        private void RemoveRecordFromLists(int recordId)
        {
            _uiContext.Post(_ =>
            {
                // Удаляем из _allRecords
                var recordToRemove = _allRecords.FirstOrDefault(r => r.Id == recordId);
                if (recordToRemove != null)
                    _allRecords.Remove(recordToRemove);

                // Удаляем из FilteredList
                var filterToRemove = FilteredList.FirstOrDefault(r => r.Id == recordId);
                if (filterToRemove != null)
                    FilteredList.Remove(filterToRemove);

                if (SelectedRecord?.Id == recordId)
                    SelectedRecord = null;

                CountStatusList = $"Всего доступно {_allRecords.Count} записей, выбрано {FilteredList.Count} записей";
            }, null);
        }

        public async Task SetRecordState(Enum status)
        {
            await SetRecordState(Convert.ToInt32(status));
        }

        public async Task SetRecordState()
        {
            await SetRecordState((int?)null);
        }

        public async Task SetRecordState(int? status)
        {
            var selectedRecord = SelectedRecord;
            if (selectedRecord == null) return;
            var article = selectedRecord.Article;

            var updatedRecord = await _notPriceComponentsService.SetRecordStateAsync(article, status)
                   .ConfigureAwait(false);
            if (updatedRecord == null)
            {
                MessageError($"Запись с артикулом {article} не найдена", "Ошибка обновления");
                return;
            }

            UpdateRecordInLists(updatedRecord);
        }

        // Обновленный BtnDeleteRecord
        public async Task BtnDeleteRecord()
        {
            if (SelectedRecord == null)
            {
                MessageWarning("Пожалуйста, выберите запись для удаления", "Запись не выбрана");
                return;
            }

            var selectedRecord = SelectedRecord;
            if (!ConfirmDelete(selectedRecord.Article)) return;

            try
            {
                bool success = await _notPriceComponentsService.DeleteRecordAsync(selectedRecord.Id)
                    .ConfigureAwait(false);

                if (success)
                {
                    // Используем общий метод удаления
                    RemoveRecordFromLists(selectedRecord.Id);
                    MessageInformation($"Запись с артикулом '{selectedRecord.Article}' удалена", "Удаление завершено");
                }
                else
                {
                    MessageWarning("Запись не была удалена", "Предупреждение");
                }
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка при удалении записи: {ex.Message}", "Ошибка удаления");
                Logger.LogException(ex);
            }
        }

        private bool ConfirmDelete(string article)
        {
            return MessageBox.Show(
                $"Вы уверены, что хотите удалить запись с артикулом '{article}'?",
                "Подтверждение удаления",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2) == DialogResult.Yes;
        }

        public async Task BtnUpdateRecord()
        {
            Range activeCell = null;
            try
            {
                activeCell = Worksheet.Application.ActiveCell;
                int currentRow = activeCell.Row;

                string article = GetCellValueAsString(Worksheet.Cells[currentRow, ArticleColumn]);

                if (string.IsNullOrWhiteSpace(article))
                {
                    MessageError("Артикул не может быть пустым", "Ошибка обновления");
                    return;
                }

                string description = GetCellValueAsString(Worksheet.Cells[currentRow, DescriptionColumn]);
                string multiplicityName = GetCellValueAsString(Worksheet.Cells[currentRow, MultiplicityColumn]);
                string productVendorName = GetCellValueAsString(Worksheet.Cells[currentRow, ProductVendorColumn]);
                decimal price = GetCellValueAsDecimal(Worksheet.Cells[currentRow, PriceColumn]);
                int discount = GetCellValueAsInt(Worksheet.Cells[currentRow, DiscountColumn]);
                string link = GetCellValueAsString(Worksheet.Cells[currentRow, LinkColumn]);

                await ProcessUpdateRecord(CreateSaveRequest(
                    article,
                    description,
                    productVendorName,
                    multiplicityName,
                    price,
                    discount,
                    link));
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка при обновлении: {ex.Message}", "Ошибка БД");
                Logger.LogException(ex);
            }
            finally
            {
                ReleaseComObjects(activeCell);
            }
        }

        private async Task ProcessUpdateRecord(NotPriceComponentSaveRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Description) || string.IsNullOrWhiteSpace(request.ProductVendorName))
            {
                MessageWarning("Описание и вендор не могут быть пустыми", "Ошибка обновления");
                return;
            }

            bool createVendorIfMissing = false;
            if (!await _notPriceComponentsService.VendorExistsAsync(request.ProductVendorName).ConfigureAwait(false))
            {
                if (!ConfirmAddNewVendor(request.ProductVendorName)) return;
                createVendorIfMissing = true;
            }

            var updatedRecord = await _notPriceComponentsService.UpdateRecordAsync(request, createVendorIfMissing).ConfigureAwait(false);
            if (updatedRecord == null)
            {
                MessageError($"Запись с артикулом {request.Article} не найдена", "Ошибка обновления");
                return;
            }

            UpdateRecordInLists(updatedRecord);

            MessageInformation($"Запись успешно обновлена\nАртикул: {updatedRecord.Article}", "Обновление завершено");
        }

        private static NotPriceComponentSaveRequest CreateSaveRequest(
            string article,
            string description,
            string productVendorName,
            string multiplicityName,
            decimal price,
            int discount,
            string link)
        {
            return new NotPriceComponentSaveRequest
            {
                Article = article?.Trim(),
                Description = description?.Trim(),
                ProductVendorName = productVendorName?.Trim(),
                MultiplicityName = multiplicityName?.Trim(),
                Price = price,
                Discount = discount,
                Link = link?.Trim()
            };
        }

        private string GetCellValueAsString(Range cell) => Convert.ToString(cell.Value2);
        private int GetCellValueAsInt(Range cell) => int.TryParse(GetCellValueAsString(cell), out int result) ? result : 0;
        private decimal GetCellValueAsDecimal(Range cell) => Convert.ToDecimal(cell.Value2);

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
