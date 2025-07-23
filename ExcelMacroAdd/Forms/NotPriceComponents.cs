using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Forms.ViewModels;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExcelMacroAdd.Forms
{
    public partial class NotPriceComponents : Form
    {
        private readonly NotPriceComponentsViewModel notPriceComponentsViewModel;
        static readonly Mutex Mutex = new Mutex(false, "MutexNotPriceComponents_SingleInstance");
        private bool _mutexAcquired = false;  

        public NotPriceComponents(INotPriceComponent accessData, IFormSettings formSettings)
        {
            notPriceComponentsViewModel = new NotPriceComponentsViewModel(accessData);
            InitializeComponent();
            InitializeDataBindings();
            SetupDataGridView();

            this.Load += async (s, e) =>
            {
                await Task.Run(() => notPriceComponentsViewModel.Start());
            };

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

            btnWritingToSheet.Click += (s, e) => notPriceComponentsViewModel.BtnWritingToSheet();
            btnAddRecord.Click += (s, e) => notPriceComponentsViewModel.BtnAddRecord();
            btnDeleteRecord.Click += (s, e) => notPriceComponentsViewModel.BtnDeleteRecord();
            btnUpdateRecord.Click += (s, e) => notPriceComponentsViewModel.BtnUpdateRecord();
            searchTextBox.TextChanged += SearchTextBox_TextChanged;
            dataGridView.SelectionChanged += DataGridView_SelectionChanged;
        }

        private void DataGridView_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView.CurrentRow != null &&
                dataGridView.CurrentRow.DataBoundItem is NotPriceComponent record)
            {
                notPriceComponentsViewModel.SelectedRecord = record;
            }
            else
            {
                notPriceComponentsViewModel.SelectedRecord = null;
            }
        }

        private void InitializeDataBindings()
        {

            // Настройка привязки данных
            dataGridView.DataSource = notPriceComponentsViewModel.FilteredList;

            // Подписка на обновления коллекции
            notPriceComponentsViewModel.PropertyChanged += (s, e) =>
            {
                // Было: nameof(NotPriceComponentsViewModel.RecordList)
                if (e.PropertyName == nameof(NotPriceComponentsViewModel.FilteredList))
                {
                    this.BeginInvoke(new Action(() =>
                    {
                        // Обновляем DataSource на FilteredList
                        dataGridView.DataSource = notPriceComponentsViewModel.FilteredList;
                    }));
                }
            };


            notPriceComponentsViewModel.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(NotPriceComponentsViewModel.CountStatusList))
                {
                    UpdateStatus(notPriceComponentsViewModel.CountStatusList);
                }
            };
        }

        private void UpdateStatus(string text)
        {
            if (statusStrip1.InvokeRequired)
            {
                statusStrip1.BeginInvoke(new Action(() => toolStripStatusLabel1.Text = text));
            }
            else
            {
                toolStripStatusLabel1.Text = text;
            }
        }


        private void SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            // Обновляем SearchTerm во ViewModel
            notPriceComponentsViewModel.SearchTerm = searchTextBox.Text;
        }

        private void SetupDataGridView()
        {             

            dataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView.ColumnHeadersDefaultCellStyle.Font =
                new System.Drawing.Font(dataGridView.Font, FontStyle.Bold);

            dataGridView.AutoSizeRowsMode =
                DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            dataGridView.ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Single;
            dataGridView.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dataGridView.GridColor = Color.Black;
            dataGridView.RowHeadersVisible = false;                                 

            dataGridView.AutoGenerateColumns = false;
            dataGridView.Columns.Clear();

            // Добавляем обработчик клика по заголовку
            dataGridView.ColumnHeaderMouseClick += DataGridView_ColumnHeaderMouseClick;

            // Настраиваем колонки для привязки данных
            // Колонка "Артикул"
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Article",
                HeaderText = "Артикул",
                Width = 110,
                SortMode = DataGridViewColumnSortMode.Programmatic // Разрешаем сортировку
            });
            // Колонка "Описание"
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Description",
                HeaderText = "Описание",
                Width = 440,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });

            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "MultiplicityDisplayName", // Используем вычисляемое свойство
                HeaderText = "Кратн.",
                Width = 60,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });

            // Колонка "Вендор" с обработкой null
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "VendorDisplayName", // Используем вычисляемое свойство
                HeaderText = "Вендор",
                Width = 70,
                SortMode = DataGridViewColumnSortMode.Programmatic // Разрешаем сортировку
            });

            // Колонка "Цена" с форматированием
            var priceColumn = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Price",
                HeaderText = "Цена",
                Width = 70,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            priceColumn.DefaultCellStyle.Format = "N2";
            priceColumn.DefaultCellStyle.NullValue = "0.00";
            dataGridView.Columns.Add(priceColumn);

            // Колонка "Скидка"
            var discountColumn = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Discount",
                HeaderText = "Скидка",
                Width = 55,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            discountColumn.DefaultCellStyle.Format = "N0";
            dataGridView.Columns.Add(discountColumn);


            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "DataRecordDisplayName", // Используем вычисляемое свойство
                HeaderText = "Дата",
                Width = 66,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });

            dataGridView.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;
            dataGridView.MultiSelect = false;

            dataGridView.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f => f.SortMode = DataGridViewColumnSortMode.NotSortable);
            dataGridView.ReadOnly = true;

            dataGridView.BackgroundColor = Color.White;                                     
        }

        private void DataGridView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var column = dataGridView.Columns[e.ColumnIndex];

            // Сортируем только разрешенные колонки
            if (column.DataPropertyName == "Article" || column.Name == "VendorDisplayName")
            {
                // Определяем направление сортировки
                ListSortDirection direction;
                if (column.SortMode == DataGridViewColumnSortMode.Programmatic &&
      (column.Name == "Article" || column.Name == "VendorDisplayName"))
                {
                    direction = ListSortDirection.Descending;
                    column.HeaderCell.SortGlyphDirection = SortOrder.Descending;
                }
                else
                {
                    direction = ListSortDirection.Ascending;

                    try { 
                        column.HeaderCell.SortGlyphDirection = SortOrder.Ascending;
                    }
                    catch (Exception ex)
                    {
                    Debug.WriteLine("!!!!!!!!!!!!!!");
                    Debug.WriteLine(ex);
                    Debug.WriteLine("!!!!!!!!!!!!!!");
                }
            }
                // Сбрасываем значки сортировки у других колонок
                //foreach (DataGridViewColumn col in dataGridView.Columns)
                //{
                //    if (col != column)
                //    {
                //        col.HeaderCell.SortGlyphDirection = SortOrder.None;
                //    }
                //}

                // Выполняем сортировку
                SortData(column.DataPropertyName, direction);
            }
        }

        private void SortData(string propertyName, ListSortDirection direction)
        {
            if (notPriceComponentsViewModel.FilteredList is BindingList<NotPriceComponent> list)
            {
                // Создаем временный список для сортировки
                var sortedList = new List<NotPriceComponent>(list);

                // Сортируем в зависимости от направления
                if (direction == ListSortDirection.Ascending)
                {
                    sortedList = propertyName == "Article"
                        ? sortedList.OrderBy(x => x.Article).ToList()
                        : sortedList.OrderBy(x => x.VendorDisplayName).ToList();
                }
                else
                {
                    sortedList = propertyName == "Article"
                        ? sortedList.OrderByDescending(x => x.Article).ToList()
                        : sortedList.OrderByDescending(x => x.VendorDisplayName).ToList();
                }

                // Обновляем FilteredList
                notPriceComponentsViewModel.FilteredList = new BindingList<NotPriceComponent>(sortedList);
            }
        }

        private void NotPriceComponents_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (_mutexAcquired)
            {
                Mutex.ReleaseMutex();
                _mutexAcquired = false;
            }
        }
    }
}
