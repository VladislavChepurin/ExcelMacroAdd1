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
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms.ViewModels
{
    public class NotPriceComponentsViewModel : AbstractFunctions, INotifyPropertyChanged
    {
        private const int ArticleColumn = 1;
        private const int DescriptionColumn = 2;
        private const int ProductVendorColumn = 5;
        private const int DiscountColumn = 6;
        private const int PriceColumn = 7;     

        private readonly INotPriceComponent accessData;
        private BindingList<NotPriceComponent> _filteredList;
        private BindingList<NotPriceComponent> _recordList;
        private NotPriceComponent _selectedRecord;
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

        public NotPriceComponent SelectedRecord
        {
            get => _selectedRecord;
            set
            {
                _selectedRecord = value;
                OnPropertyChanged(nameof(SelectedRecord));
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
            
            if (string.IsNullOrWhiteSpace(article))
            {
                MessageError($"Добавить невозможно, пустой артикул", "Ошибка добавления");
                return;
            }

            if (await accessData.AccessNotPriceComponent.IsThereIsDBRecord(article))
            {
                MessageError($"Артикул {article} уже есть в базе дданных  добавить невозможно", "Ошибка добавления");
                return;
            }

            int.TryParse(Convert.ToString(Worksheet.Cells[currentRow, DiscountColumn].Value2), out int discount);
            string description = Convert.ToString(Worksheet.Cells[currentRow, DescriptionColumn].Value2);
            string productVendorName = Convert.ToString(Worksheet.Cells[currentRow, ProductVendorColumn].Value2);
            double price = Convert.ToDouble(Worksheet.Cells[currentRow, PriceColumn].Value2);     

            if (string.IsNullOrEmpty(article) || string.IsNullOrEmpty(description) || string.IsNullOrEmpty(productVendorName))
            {
                MessageWarning($"Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \n Артикул = {article}",
                    "Ошибка записи");     
                return;
            }

            var productVendorEntity = await accessData.AccessNotPriceComponent.GetProductVendorEntityByName(productVendorName);
            // Если вендора нет в базе
            if (productVendorEntity == null)
            {
                // Запрос пользователю
                var dialogResult = MessageBox.Show(
                    $"В БД вендора '{productVendorName}' нет. Добавить нового вендора?",
                    "Добавление нового вендора",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (dialogResult != DialogResult.Yes)
                {
                    MessageInformation("Добавление записи отменено", "Отмена операции");
                    return;
                }

                // Создаем нового вендора                
                productVendorEntity = await accessData.AccessNotPriceComponent.AddProductVendor(new ProductVendor { VendorName = productVendorName });              
            }

            // Создаем и сохраняем запись
            var entity = new NotPriceComponent()
            {
                Article = article,
                Description = description,
                ProductVendorId = productVendorEntity.Id,
                Price = price,
                Discount = discount
            };

            await accessData.AccessNotPriceComponent.AddValueDb(entity);
            // Обновляем данные
            Start();

            MessageInformation($"Успешно записано в базу данных!\nАртикул: {article}\nВендор: {productVendorName}",
                "Запись успешна!");        
        }

        public async void BtnDeleteRecord()
        {
            if (SelectedRecord == null)
            {
                MessageWarning("Пожалуйста, выберите запись для удаления", "Запись не выбрана");
                return;
            }

            var selectedRecord = SelectedRecord;
            var article = selectedRecord.Article;

            // Запрос подтверждения
            var dialogResult = MessageBox.Show(
                $"Вы уверены, что хотите удалить запись с артикулом '{article}'?",
                "Подтверждение удаления",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);

            if (dialogResult != DialogResult.Yes)
                return;

            try
            {
                // Выполняем удаление
                bool success = await accessData.AccessNotPriceComponent.DeleteRecord(selectedRecord.Id);

                if (success)
                {
                    // Удаляем запись из коллекции
                    RecordList.Remove(selectedRecord);

                    // Сбрасываем выделение
                    SelectedRecord = null;

                    MessageInformation($"Запись с артикулом '{article}' успешно удалена", "Удаление завершено");
                }
                else
                {
                    MessageWarning("Запись не была удалена. Возможно, она уже была удалена ранее.", "Предупреждение");
                }
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка при удалении записи: {ex.Message}", "Ошибка удаления");
            }
        }


        public async void BtnUpdateRecord()
        {
            var currentRow = Cell.Row; // Текущая строка в Excel
            string article = Convert.ToString(Worksheet.Cells[currentRow, ArticleColumn].Value2);

            if (string.IsNullOrWhiteSpace(article))
            {
                MessageError("Артикул не может быть пустым", "Ошибка обновления");
                return;
            }

            try
            {
                // Ищем запись по артикулу
                var existingRecord = await accessData.AccessNotPriceComponent.GetRecordByArticle(article);
                if (existingRecord == null)
                {
                    MessageError($"Запись с артикулом {article} не найдена", "Ошибка обновления");
                    return;
                }

                // Получаем новые данные из Excel
                string description = Convert.ToString(Worksheet.Cells[currentRow, DescriptionColumn].Value2);
                string productVendorName = Convert.ToString(Worksheet.Cells[currentRow, ProductVendorColumn].Value2);
                double price = Convert.ToDouble(Worksheet.Cells[currentRow, PriceColumn].Value2);
                int.TryParse(Convert.ToString(Worksheet.Cells[currentRow, DiscountColumn].Value2), out int discount);

                // Проверка обязательных полей
                if (string.IsNullOrWhiteSpace(description) || string.IsNullOrWhiteSpace(productVendorName))
                {
                    MessageWarning("Описание и вендор не могут быть пустыми", "Ошибка обновления");
                    return;
                }

                // Обновляем данные
                existingRecord.Description = description;
                existingRecord.Price = (float)price;
                existingRecord.Discount = discount;

                // Обработка вендора
                
                
                var productVendorEntity = await accessData.AccessNotPriceComponent.GetProductVendorEntityByName(productVendorName);
                // Если вендора нет в базе
                if (productVendorEntity == null)
                {
                    var dialogResult = MessageBox.Show(
                        $"В БД вендора '{productVendorName}' нет. Добавить нового вендора?",
                        "Добавление нового вендора",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (dialogResult != DialogResult.Yes)
                    {
                        MessageInformation("Обновление записи отменено", "Отмена операции");
                        return;
                    }

                    // Создаем нового вендора                
                    productVendorEntity = await accessData.AccessNotPriceComponent.AddProductVendor(new ProductVendor { VendorName = productVendorName });
                                       
                    existingRecord.ProductVendorId = productVendorEntity.Id;
                }
                else
                {
                    existingRecord.ProductVendorId = productVendorEntity.Id;
                }

                // Сохраняем изменения
                await accessData.AccessNotPriceComponent.UpdateRecord(existingRecord);

                // Обновляем данные
                Start();

                MessageInformation($"Запись успешно обновлена\nАртикул: {article}", "Обновление завершено");
            }
            catch (Exception ex)
            {
                MessageError($"Ошибка при обновлении: {ex.Message}", "Ошибка БД");
            }
        }



        // Реализация INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}