using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Services;
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

        private readonly INotPriceComponent _accessData;
        private BindingList<NotPriceComponent> _filteredList;
        private BindingList<NotPriceComponent> _recordList;
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

        public BindingList<NotPriceComponent> RecordList
        {
            get => _recordList;
            set
            {
                if (_recordList != value)
                {
                    _recordList = value;
                    OnPropertyChanged(nameof(RecordList));
                }
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

        public NotPriceComponentsViewModel(INotPriceComponent accessData)
        {
            _accessData = accessData ?? throw new ArgumentNullException(nameof(accessData));
            _filterTokenSource = new CancellationTokenSource();
            FilteredList = new BindingList<NotPriceComponent>();
            _uiContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();
        }

        public override async void Start()
        {
            try
            {
                IsLoading = true;
                var records = await _accessData.AccessNotPriceComponent.GetAllRecord().ConfigureAwait(false);
                RecordList = new BindingList<NotPriceComponent>(records.ToList());
                FilteredList = new BindingList<NotPriceComponent>(records.ToList());
                CountStatusList = $"Всего доступно {RecordList.Count} записей, выбрано {FilteredList.Count} записей";
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
                IEnumerable<NotPriceComponent> result;

                if (string.IsNullOrWhiteSpace(search))
                {
                    result = RecordList;
                }
                else
                {
                    result = RecordList
                        .AsParallel()
                        .WithCancellation(token)
                        .Where(item =>
                            item != null &&
                            (!string.IsNullOrEmpty(item.Article) &&
                             item.Article.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (!string.IsNullOrEmpty(item.Description) &&
                             item.Description.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (!string.IsNullOrEmpty(item.VendorDisplayName) &&
                             item.VendorDisplayName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0))
                        .AsEnumerable(); // Преобразуем обратно в IEnumerable

                    if (result.Count() > MaxDisplayItems)
                    {
                        result = result.Take(MaxDisplayItems);
                    }
                }

                FilteredList = new BindingList<NotPriceComponent>(result.ToList());

                CountStatusList = $"Всего доступно {RecordList.Count} записей, выбрано {FilteredList.Count} записей";
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

            var activeCell = Worksheet.Application.ActiveCell;
            int currentRow = activeCell.Row;

            try
            {              
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
                    Marshal.FinalReleaseComObject(obj);
                }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public async void BtnAddRecord()
        {
            var activeCell = Worksheet.Application.ActiveCell;
            int currentRow = activeCell.Row;
            try
            {              
                string article = GetCellValueAsString(Worksheet.Cells[currentRow, ArticleColumn]);

                if (string.IsNullOrWhiteSpace(article))
                {
                    MessageError("Добавить невозможно, пустой артикул", "Ошибка добавления");
                    return;
                }

                if (await _accessData.AccessNotPriceComponent.IsThereIsDBRecord(article).ConfigureAwait(false))
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

                await ProcessAddRecord(article, description, productVendorName, multiplicityName, price, discount, link);
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

        private async Task ProcessAddRecord(string article, string description, string productVendorName,
                                          string multiplicityName, decimal price, int discount, string link)
        {
            var productVendorEntity = await _accessData.AccessNotPriceComponent.GetProductVendorEntityByName(productVendorName)
                .ConfigureAwait(false);

            if (productVendorEntity == null)
            {
                if (!ConfirmAddNewVendor(productVendorName)) return;
                productVendorEntity = await _accessData.AccessNotPriceComponent.AddProductVendor(
                    new ProductVendor { VendorName = productVendorName }).ConfigureAwait(false);
            }

            var multiplicityEntity = await _accessData.AccessNotPriceComponent.GetMultiplicityEntityByName(multiplicityName)
                .ConfigureAwait(false) ?? new Multiplicity() { Id = 1 };

            var entity = new NotPriceComponent
            {         
                Article = article,
                Description = description,
                MultiplicityId = multiplicityEntity.Id,
                ProductVendorId = productVendorEntity.Id,
                Price = price,
                Discount = discount,
                DataRecord = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"),
                Link = link
            };

            await _accessData.AccessNotPriceComponent.AddValueDb(entity).ConfigureAwait(false);
            Start();

            MessageInformation($"Успешно записано в базу данных!\nАртикул: {article}\nВендор: {productVendorName}",
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
                // Обновляем в RecordList
                var recordItem = RecordList.FirstOrDefault(r => r.Id == updatedRecord.Id);
                if (recordItem != null)
                {
                    var index = RecordList.IndexOf(recordItem);
                    RecordList[index] = updatedRecord;
                }

                // Обновляем в FilteredList
                var filteredItem = FilteredList.FirstOrDefault(r => r.Id == updatedRecord.Id);
                if (filteredItem != null)
                {
                    var index = FilteredList.IndexOf(filteredItem);
                    FilteredList[index] = updatedRecord;
                }

                CountStatusList = $"Всего доступно {RecordList.Count} записей, выбрано {FilteredList.Count} записей";
            }, null);
        }

        // Общий метод для удаления записи
        private void RemoveRecordFromLists(int recordId)
        {
            _uiContext.Post(_ =>
            {
                // Удаляем из RecordList
                var recordToRemove = RecordList.FirstOrDefault(r => r.Id == recordId);
                if (recordToRemove != null)
                    RecordList.Remove(recordToRemove);

                // Удаляем из FilteredList
                var filterToRemove = FilteredList.FirstOrDefault(r => r.Id == recordId);
                if (filterToRemove != null)
                    FilteredList.Remove(filterToRemove);

                if (SelectedRecord?.Id == recordId)
                    SelectedRecord = null;

                CountStatusList = $"Всего доступно {RecordList.Count} записей, выбрано {FilteredList.Count} записей";
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

            var existingRecord = await _accessData.AccessNotPriceComponent.GetRecordByArticle(article)
                   .ConfigureAwait(false);
            if (existingRecord == null)
            {
                MessageError($"Запись с артикулом {article} не найдена", "Ошибка обновления");
                return;
            }

            existingRecord.IsValid = status;
            await _accessData.AccessNotPriceComponent.UpdateRecord(existingRecord).ConfigureAwait(false);

            // Используем общий метод обновления
            UpdateRecordInLists(existingRecord);
        }

        // Обновленный BtnDeleteRecord
        public async void BtnDeleteRecord()
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
                bool success = await _accessData.AccessNotPriceComponent.DeleteRecord(selectedRecord.Id)
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

        public async void BtnUpdateRecord()
        {
            var activeCell = Worksheet.Application.ActiveCell;
            int currentRow = activeCell.Row;

            try
            {                
                string article = GetCellValueAsString(Worksheet.Cells[currentRow, ArticleColumn]);

                if (string.IsNullOrWhiteSpace(article))
                {
                    MessageError("Артикул не может быть пустым", "Ошибка обновления");
                    return;
                }

                var existingRecord = await _accessData.AccessNotPriceComponent.GetRecordByArticle(article)
                    .ConfigureAwait(false);

                if (existingRecord == null)
                {
                    MessageError($"Запись с артикулом {article} не найдена", "Ошибка обновления");
                    return;
                }

                await ProcessUpdateRecord(currentRow, existingRecord);
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

        private async Task ProcessUpdateRecord(int currentRow, NotPriceComponent existingRecord)
        {
            string description = GetCellValueAsString(Worksheet.Cells[currentRow, DescriptionColumn]);
            string multiplicityName = GetCellValueAsString(Worksheet.Cells[currentRow, MultiplicityColumn]);
            string productVendorName = GetCellValueAsString(Worksheet.Cells[currentRow, ProductVendorColumn]);
            decimal price = GetCellValueAsDecimal(Worksheet.Cells[currentRow, PriceColumn]);
            int discount = GetCellValueAsInt(Worksheet.Cells[currentRow, DiscountColumn]);
            string link = GetCellValueAsString(Worksheet.Cells[currentRow, LinkColumn]);

            if (string.IsNullOrWhiteSpace(description) || string.IsNullOrWhiteSpace(productVendorName))
            {
                MessageWarning("Описание и вендор не могут быть пустыми", "Ошибка обновления");
                return;
            }

            existingRecord.Description = description;
            existingRecord.Price = price;                   
            existingRecord.Discount = discount;
            existingRecord.DataRecord = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss");
            if (!string.IsNullOrWhiteSpace(link))
                existingRecord.Link = link;

            var productVendorEntity = await _accessData.AccessNotPriceComponent.GetProductVendorEntityByName(productVendorName)
                .ConfigureAwait(false);

            if (productVendorEntity == null)
            {
                if (!ConfirmAddNewVendor(productVendorName)) return;
                productVendorEntity = await _accessData.AccessNotPriceComponent.AddProductVendor(
                    new ProductVendor { VendorName = productVendorName }).ConfigureAwait(false);
            }

            var multiplicityEntity = await _accessData.AccessNotPriceComponent.GetMultiplicityEntityByName(multiplicityName)
                .ConfigureAwait(false) ?? new Multiplicity() { Id = 1 };

            existingRecord.ProductVendorId = productVendorEntity.Id;
            existingRecord.MultiplicityId = multiplicityEntity.Id;

            await _accessData.AccessNotPriceComponent.UpdateRecord(existingRecord).ConfigureAwait(false);
            //Обнавление измеенной записи
            UpdateRecordInLists(existingRecord);

            MessageInformation($"Запись успешно обновлена\nАртикул: {existingRecord.Article}", "Обновление завершено");
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