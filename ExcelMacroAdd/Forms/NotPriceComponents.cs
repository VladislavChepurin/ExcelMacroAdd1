using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Forms.CustomUI;
using ExcelMacroAdd.Forms.ViewModels;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
        private ListSortDirection _currentSortDirection = ListSortDirection.Ascending;
        private string _currentSortProperty = "Article";

        enum status
        {
            check,
            discontinued,
            question
        }

        public NotPriceComponents(INotPriceComponent accessData, IFormSettings formSettings)
        {
            notPriceComponentsViewModel = new NotPriceComponentsViewModel(accessData);            
            InitializeComponent();
            InitializeDataBindings();
            SetupDataGridView();
            SetupContextMenu();

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
            linkToTheWebsite.Click += (s, e) => notPriceComponentsViewModel.OpenLink();
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

            // Привязка для ссылки
            linkToTheWebsite.DataBindings.Add("Text",
                notPriceComponentsViewModel,
                nameof(NotPriceComponentsViewModel.DisplayLink),
                false,
                DataSourceUpdateMode.OnPropertyChanged);

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
            //dataGridView.EnableHeadersVisualStyles = false;
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
            
            var isValidColumn = new DataGridViewImageColumn
            {
                HeaderText = "",
                Width = 20,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                ImageLayout = DataGridViewImageCellLayout.Zoom
            };
            dataGridView.Columns.Add(isValidColumn);

            // ДОБАВЛЯЕМ ОБРАБОТЧИК ЗДЕСЬ
            dataGridView.DataBindingComplete += DataGridView_DataBindingComplete;

            // Колонка "Артикул"
            var articleColumn = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Article",
                HeaderText = "Артикул",
                Width = 110,
                SortMode = DataGridViewColumnSortMode.Programmatic
            };
            articleColumn.HeaderCell = new CustomDataGridViewHeaderCell("Артикул");
            dataGridView.Columns.Add(articleColumn);

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
            var vendorColumn = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "VendorDisplayName",
                HeaderText = "Вендор",
                Width = 70,
                SortMode = DataGridViewColumnSortMode.Programmatic
            };
            vendorColumn.HeaderCell = new CustomDataGridViewHeaderCell("Вендор");
            dataGridView.Columns.Add(vendorColumn);

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

            dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView.MultiSelect = false;
           
            dataGridView.ReadOnly = true;
            dataGridView.BackgroundColor = Color.White;                                     
        }

        private void DataGridView_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.DataBoundItem is NotPriceComponent component)
                {
                    var cell = row.Cells[0]; // колонка со статусом
                    if (component.IsValid == 0)
                        cell.Value = Properties.Resources.check;
                    else if (component.IsValid == 1)
                        cell.Value = Properties.Resources.delete;
                    else if (component.IsValid == 2)
                        cell.Value = Properties.Resources.question;
                    else
                        cell.Value = Properties.Resources.noneImage;
                }
            }
        }
                
        private void SetupContextMenu()
        {
            // Создаем контекстное меню
            ContextMenuStrip contextMenu = new ContextMenuStrip();

            // Добавляем пункты меню            
            contextMenu.Items.Add("На лист", null, (s, e) => notPriceComponentsViewModel.BtnWritingToSheet());

            contextMenu.Items.Add("Копировать артикул", null, (s, e) =>
            {
                Task.Run(() =>
                {
                    // Создаем STA поток для работы с буфером обмена
                    var thread = new Thread(() =>
                    {
                        CopySelectedArticle();
                    });
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
                });
            });
            contextMenu.Items.Add("Копировать описание", null, (s, e) =>
            {
                Task.Run(() =>
                {
                    // Создаем STA поток для работы с буфером обмена
                    var thread = new Thread(() =>
                    {
                        CopySelectedDescription();
                    });
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
                });
            });
            contextMenu.Items.Add("Удалить запись", null, (s, e) => 
            notPriceComponentsViewModel.BtnDeleteRecord());
            contextMenu.Items.Add(new ToolStripSeparator());
            contextMenu.Items.Add("Проверено", null, async (s, e) =>
            {
                await notPriceComponentsViewModel.SetRecordState(status.check);
            });
            contextMenu.Items.Add("Снято с производства", null, async (s, e) =>
            {
                await notPriceComponentsViewModel.SetRecordState(status.discontinued);
            });
            contextMenu.Items.Add("Под сомнением", null, async (s, e) =>
            {
                await notPriceComponentsViewModel.SetRecordState(status.question);
            });
            contextMenu.Items.Add("Сбросить поле", null, async (s, e) =>
            {
                await notPriceComponentsViewModel.SetRecordState();
            });

            // Привязываем меню к DataGridView (ОБЯЗАТЕЛЬНО ДО MouseDown)
            dataGridView.ContextMenuStrip = contextMenu;

            // Обработчик MouseDown для выделения строки правой кнопкой
            dataGridView.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    var hitTest = dataGridView.HitTest(e.X, e.Y);
                    if (hitTest.RowIndex >= 0)
                    {
                        // Устанавливаем CurrentCell, чтобы сработал SelectionChanged
                        if (hitTest.ColumnIndex >= 0)
                        {
                            dataGridView.CurrentCell = dataGridView.Rows[hitTest.RowIndex].Cells[hitTest.ColumnIndex];
                        }

                        // Выделяем строку
                        dataGridView.ClearSelection();
                        dataGridView.Rows[hitTest.RowIndex].Selected = true;

                        // Дополнительно: явно вызываем обработчик, если нужно
                        // DataGridView_SelectionChanged(null, null);
                    }
                }
            };

            // Обработчик открытия меню
            contextMenu.Opening += (s, e) =>
            {
                // Получаем позицию курсора относительно DataGridView
                Point mousePosition = dataGridView.PointToClient(Cursor.Position);
                var hitTest = dataGridView.HitTest(mousePosition.X, mousePosition.Y);

                // Проверяем все условия для открытия меню
                if (hitTest.RowIndex >= 0 && dataGridView.SelectedRows.Count > 0)
                {
                    e.Cancel = false; // Разрешаем открытие
                }
                else
                {
                    e.Cancel = true;  // Запрещаем открытие
                }
            };
        }
        
        private void CopySelectedArticle()
        {
            if (dataGridView.SelectedRows.Count > 0)
            {
                if (dataGridView.SelectedRows[0].DataBoundItem is NotPriceComponent component && !string.IsNullOrEmpty(component.Article))
                {
                    Clipboard.SetText(component.Article);
                    MessageBox.Show($"Артикул '{component.Article}' скопирован",
                        "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        private void CopySelectedDescription()
        {
            if (dataGridView.SelectedRows.Count > 0)
            {
                if (dataGridView.SelectedRows[0].DataBoundItem is NotPriceComponent component && !string.IsNullOrEmpty(component.Description))
                {
                    Clipboard.SetText(component.Description);
                    MessageBox.Show($"Описание '{component.Description}' скопировано",
                        "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void DataGridView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var column = dataGridView.Columns[e.ColumnIndex];

            if (column.DataPropertyName == "Article" || column.DataPropertyName == "VendorDisplayName")
            {
                // Определяем новое направление сортировки
                ListSortDirection newDirection;

                if (column.DataPropertyName == _currentSortProperty)
                {
                    newDirection = _currentSortDirection == ListSortDirection.Ascending
                        ? ListSortDirection.Descending
                        : ListSortDirection.Ascending;
                }
                else
                {
                    _currentSortProperty = column.DataPropertyName;
                    newDirection = ListSortDirection.Ascending;
                }

                _currentSortDirection = newDirection;

                // Сбрасываем значки у всех колонок
                foreach (DataGridViewColumn col in dataGridView.Columns)
                {
                    if (col.HeaderCell is CustomDataGridViewHeaderCell headerCell)
                    {
                        headerCell.SortGlyphDirection = SortOrder.None;
                    }
                }

                // Устанавливаем значок для текущей колонки
                if (column.HeaderCell is CustomDataGridViewHeaderCell currentHeader)
                {
                    currentHeader.SortGlyphDirection = newDirection == ListSortDirection.Ascending
                        ? SortOrder.Ascending
                        : SortOrder.Descending;
                }

                // Выполняем сортировку
                SortData(column.DataPropertyName, newDirection);
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
