using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Functions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelMacroAdd.Forms.ViewModels
{
    public class NotPriceComponentsViewModel : AbstractFunctions, INotifyPropertyChanged
    {
        private const int ArticleColumn = 1;
        private const int DescriptionColumn = 2;
        private const int ProductVendorColumn = 5;
        private const int PriceColumn = 6;
        private const int DiscountColumn = 7;

        private readonly INotPriceComponent accessData;
        private BindingList<NotPriceComponent> _filteredList;
        private BindingList<NotPriceComponent> _recordList;
        private string _searchTerm;
        private CancellationTokenSource _filterTokenSource;
        private bool _isLoading;

        public BindingList<NotPriceComponent> RecordList
        {
            get => _recordList;
            set
            {
                _recordList = value;
                OnPropertyChanged(nameof(RecordList));
            }
        }



        public BindingList<NotPriceComponent> FilteredList
        {
            get => _filteredList;
            set
            {
                _filteredList = value;
                OnPropertyChanged(nameof(FilteredList));
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
                    ApplyFilterAsync(); // Фильтруем при изменении текста
                }
            }
        }

        public bool IsLoading
        {
            get => _isLoading;
            set
            {
                _isLoading = value;
                OnPropertyChanged(nameof(IsLoading));
            }
        }

        public NotPriceComponentsViewModel(INotPriceComponent accessData)
        {
            this.accessData = accessData;
            _filterTokenSource = new CancellationTokenSource();
            FilteredList = new BindingList<NotPriceComponent>();
        }

        public override async void Start()
        {
            try
            {
                IsLoading = true;
                var records = await accessData.AccessNotPriceComponent.GetAllRecord();
                RecordList = new BindingList<NotPriceComponent>(records.ToList());
                FilteredList = RecordList;
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка загрузки данных: {ex.Message}", "Ошибка загрузки");             
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async void ApplyFilterAsync()
        {
            // Отменяем предыдущую операцию фильтрации
            _filterTokenSource.Cancel();
            _filterTokenSource = new CancellationTokenSource();
            var token = _filterTokenSource.Token;

            try
            {
                // Задержка для оптимизации ввода
                await Task.Delay(300, token);
                if (token.IsCancellationRequested) return;
                              
                var search = SearchTerm?.Trim();              

                // Фильтрация с использованием параллельной обработки
                List<NotPriceComponent> result;
                if (string.IsNullOrWhiteSpace(search))
                {
                    result = RecordList.ToList();
                }
                else
                {
                    result = RecordList
                        .AsParallel()
                        .WithCancellation(token)
                        .Where(item =>
                            (item.Article != null && item.Article.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (item.Description != null && item.Description.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (item.VendorDisplayName != null && item.VendorDisplayName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0))
                        .ToList();
                }

                // Ограничение количества отображаемых элементов
                if (result.Count > 1000)
                {
                    result = result.Take(1000).ToList();
                }

                FilteredList = new BindingList<NotPriceComponent>(result);                
            }
            catch (TaskCanceledException)
            {
                // Фильтрация была отменена - ничего не делаем
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка фильтрации: {ex.Message}");
            }
        }

        public void BtnWritingToSheet()
        {
            throw new NotImplementedException();
        }

        public async void BtnAddRecord()
        {
            var currentRow = Cell.Row; // Вычисляем верхний элемент
            string article = Convert.ToString(Worksheet.Cells[currentRow, ArticleColumn].Value2);
            
            if (article is null)
            {
                MessageError($"Добавить невозможно, пустой артикул", "Ошибка добавления");
                return;
            }

            if (await accessData.AccessNotPriceComponent.IsThereIsDBRecord(article))
            {
                MessageError($"Артикул {article} уже есть в базе дданных  добавить невозможно", "Ошибка добавления");
                return;
            }

            int.TryParse(Convert.ToString(Worksheet.Cells[currentRow, IPRatingColumn].Value2), out int discount);

            string description = Convert.ToString(Worksheet.Cells[currentRow, DescriptionColumn].Value2);
            string productVendor = Convert.ToString(Worksheet.Cells[currentRow, ProductVendorColumn].Value2);
            double price = Convert.ToDouble(Worksheet.Cells[currentRow, PriceColumn].Value2);     

            if (string.IsNullOrEmpty(article) || string.IsNullOrEmpty(description) || string.IsNullOrEmpty(productVendor))
            {
                MessageWarning($"Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \n Артикул = {article}",
                    "Ошибка записи");     
                return;
            }

            var productVendorEntity = await accessData.AccessNotPriceComponent.GetProductVendorEntityByName(productVendor);

            var etity = new NotPriceComponent()
            {
                Article = article,
                Description = description,
                ProductVendorId = productVendorEntity.Id,
                Price = (float)price,
                Discount = discount
            };

            await accessData.AccessNotPriceComponent.AddValueDb(etity);

            MessageInformation($"Успешно записано в базу данных. Теперь доступна новая запись.\n Поздравляем! \nАртикул = {article}",
                       "Запись успешна!");
        }

        public void BtnDeleteRecord()
        {
            throw new NotImplementedException();
        }

        // Реализация INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}