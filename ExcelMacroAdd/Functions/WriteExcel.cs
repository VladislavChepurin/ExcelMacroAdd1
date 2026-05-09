using ExcelMacroAdd.Services;
using ExcelMacroAdd.Services.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace ExcelMacroAdd.Functions
{
    internal sealed class WriteExcel : AbstractFunctions
    {
        private readonly IDataInXml _dataInXml;
        private readonly string _vendor;
        private readonly int _startRow;
        private readonly string _article;
        private readonly int _amount;
        private readonly int _countRows;

        private const int ArticleColumn = 1;
        private const int DescriptionColumn = 2;
        private const int QuantityColumn = 3;
        private const int MultiplicityColumn = 4;
        private const int VendorColumn = 5;
        private const int DiscountColumn = 6;
        private const int PriceColumn = 7;
        private const int TotalPriceColumn = 8;
        private const int CoastColumn = 9;
        private const int DateColumn = 10;

        public WriteExcel(IDataInXml dataInXml, string vendor)
        {
            _dataInXml = dataInXml ?? throw new ArgumentNullException(nameof(dataInXml));
            _vendor = vendor ?? throw new ArgumentNullException(nameof(vendor));

            Range activeCell = null;
            Range selectedRows = null;

            try
            {
                activeCell = Cell;
                selectedRows = activeCell.Rows;
                _startRow = activeCell.Row;
                _countRows = selectedRows.Count;
            }
            finally
            {
                ReleaseComObject(selectedRows);
                ReleaseComObject(activeCell);
            }
        }

        public WriteExcel(IDataInXml dataInXml, string vendor, string article, int startOffset = 0, int amount = 0)
        {
            _dataInXml = dataInXml ?? throw new ArgumentNullException(nameof(dataInXml));
            _vendor = vendor ?? throw new ArgumentNullException(nameof(vendor));
            _article = article;
            _amount = amount;
            _countRows = 1;

            Range activeCell = null;

            try
            {
                activeCell = Cell;
                _startRow = activeCell.Row + startOffset;
            }
            finally
            {
                ReleaseComObject(activeCell);
            }
        }

        public override void Start()
        {
            if (ExcelPerformanceScope.CurrentNestingLevel == 0)
            {
                throw new InvalidOperationException(
                    "WriteExcel.Start() must be called inside ExcelPerformanceScope");
            }

            Worksheet worksheet = null;

            try
            {
                worksheet = Worksheet;
                if (worksheet == null)
                {
                    throw new InvalidOperationException("Не инициализирован объект Excel.");
                }

                var vendors = _dataInXml.ReadFileXml();
                var vendorData = _dataInXml.ReadElementXml(_vendor, vendors)
                    ?? throw new ArgumentException($"Вендор {_vendor} не найден.");

                if (_countRows <= 1)
                {
                    WriteSingleRow(worksheet, _startRow, vendorData);
                }
                else
                {
                    WriteBatchRows(worksheet, vendorData);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
            }
            finally
            {
                ReleaseComObject(worksheet);
            }
        }

        /// <summary>
        /// Запись одной строки.
        /// Порядок: сначала все лёгкие значения, формулы ВПР — в самом конце.
        /// </summary>
        private void WriteSingleRow(Worksheet worksheet, int row, UserVariables.Vendor vendorData)
        {
            // 1) Лёгкие значения — мгновенно
            if (_article != null)
            {
                SetCellValue(worksheet, row, ArticleColumn, _article);
            }

            if (_amount != 0)
            {
                SetCellValue(worksheet, row, QuantityColumn, _amount);
            }

            SetCellValue(worksheet, row, VendorColumn, _vendor);
            SetCellValue(worksheet, row, DiscountColumn, vendorData.Discount);
            SetCellNumberFormat(worksheet, row, DateColumn, "ДД.ММ.ГГ ч:мм");
            SetCellValue(worksheet, row, DateColumn, DateTime.Now);

            // 2) Простые формулы (не ВПР) — быстрые
            SetCellFormula(worksheet, row, TotalPriceColumn, $"=G{row}*(100-F{row})/100");
            SetCellFormula(worksheet, row, CoastColumn, $"=H{row}*C{row}");

            // 3) Формулы ВПР — тяжёлые, в конце
            SetCellFormulaLocal(worksheet, row, DescriptionColumn, string.Format(vendorData.Formula_1, row));
            SetCellFormulaLocal(worksheet, row, MultiplicityColumn, string.Format(vendorData.Formula_2, row));
            SetCellFormulaLocal(worksheet, row, PriceColumn, string.Format(vendorData.Formula_3, row));
        }

        /// <summary>
        /// Пакетная запись нескольких строк.
        ///
        /// Стратегия «текст → формула»:
        ///   Шаг 1: Записываем ВПР-формулы как ТЕКСТ (NumberFormat = "@")
        ///          — Excel не парсит формулу, не трогает прайс, это мгновенно
        ///   Шаг 2: Снимаем текстовый формат и конвертируем в формулы
        ///          одной операцией FormulaLocal = values
        ///
        /// Почему это быстрее прямой записи FormulaLocal:
        ///   При FormulaLocal на каждую ячейку Excel парсит ВПР, резолвит
        ///   ссылку на лист прайса, ищет диапазон. С "@" запись идёт
        ///   как Value2 текстовой строки — без парсинга.
        ///   Конвертация в конце — одна операция на весь столбец.
        /// </summary>
        private void WriteBatchRows(Worksheet worksheet, UserVariables.Vendor vendorData)
        {
            var dateNow = DateTime.Now;

            // === Обычные значения — массивами ===
            var vendorValues = new object[_countRows, 1];
            var discountValues = new object[_countRows, 1];
            var dateValues = new object[_countRows, 1];

            for (int i = 0; i < _countRows; i++)
            {
                vendorValues[i, 0] = _vendor;
                discountValues[i, 0] = vendorData.Discount;
                dateValues[i, 0] = dateNow;
            }

            SetRangeValue(worksheet, VendorColumn, vendorValues);
            SetRangeValue(worksheet, DiscountColumn, discountValues);

            Range dateRange = null;
            try
            {
                dateRange = GetColumnRange(worksheet, DateColumn);
                dateRange.NumberFormat = "ДД.ММ.ГГ ч:мм";
                dateRange.Value2 = dateValues;
            }
            finally
            {
                ReleaseComObject(dateRange);
            }

            // === Простые формулы (не ВПР) ===
            var totalPriceFormulas = new object[_countRows, 1];
            var coastFormulas = new object[_countRows, 1];

            for (int i = 0; i < _countRows; i++)
            {
                int row = _startRow + i;
                totalPriceFormulas[i, 0] = $"=G{row}*(100-F{row})/100";
                coastFormulas[i, 0] = $"=H{row}*C{row}";
            }

            SetRangeFormula(worksheet, TotalPriceColumn, totalPriceFormulas);
            SetRangeFormula(worksheet, CoastColumn, coastFormulas);

            // === Формулы ВПР — двухфазная вставка ===
            var formula1Text = new object[_countRows, 1];
            var formula2Text = new object[_countRows, 1];
            var formula3Text = new object[_countRows, 1];

            for (int i = 0; i < _countRows; i++)
            {
                int row = _startRow + i;
                formula1Text[i, 0] = string.Format(vendorData.Formula_1, row);
                formula2Text[i, 0] = string.Format(vendorData.Formula_2, row);
                formula3Text[i, 0] = string.Format(vendorData.Formula_3, row);
            }

            // Фаза 1: пишем как текст (мгновенно)
            WriteAsText(worksheet, DescriptionColumn, formula1Text);
            WriteAsText(worksheet, MultiplicityColumn, formula2Text);
            WriteAsText(worksheet, PriceColumn, formula3Text);

            // Фаза 2: конвертируем текст → формулы (одна операция на столбец)
            ActivateFormulas(worksheet, DescriptionColumn);
            ActivateFormulas(worksheet, MultiplicityColumn);
            ActivateFormulas(worksheet, PriceColumn);
        }

        /// <summary>
        /// Записывает строки-формулы как текст.
        /// NumberFormat = "@" заставляет Excel трактовать "=ВПР(...)" как литерал.
        /// Value2 со строками в текстовом формате — мгновенная операция.
        /// </summary>
        private void WriteAsText(Worksheet worksheet, int column, object[,] formulaTexts)
        {
            Range range = null;

            try
            {
                range = GetColumnRange(worksheet, column);
                range.NumberFormat = "@";
                range.Value2 = formulaTexts;
            }
            finally
            {
                ReleaseComObject(range);
            }
        }

        /// <summary>
        /// Конвертирует текстовые строки формул в реальные формулы Excel.
        /// Снимает текстовый формат, читает значения и записывает обратно
        /// как FormulaLocal — одна COM-операция на весь диапазон.
        /// </summary>
        private void ActivateFormulas(Worksheet worksheet, int column)
        {
            Range range = null;

            try
            {
                range = GetColumnRange(worksheet, column);
                range.NumberFormat = "General";

                if (_countRows == 1)
                {
                    string text = range.Value2?.ToString();
                    if (!string.IsNullOrEmpty(text) && text.StartsWith("="))
                    {
                        range.FormulaLocal = text;
                    }
                }
                else
                {
                    object[,] values = range.Value2 as object[,];
                    if (values != null)
                    {
                        range.FormulaLocal = values;
                    }
                }
            }
            finally
            {
                ReleaseComObject(range);
            }
        }

        // === Вспомогательные методы ===

        private Range GetColumnRange(Worksheet worksheet, int column)
        {
            Range startCell = null;
            Range endCell = null;

            try
            {
                int endRow = _startRow + _countRows - 1;
                startCell = (Range)worksheet.Cells[_startRow, column];
                endCell = (Range)worksheet.Cells[endRow, column];
                return worksheet.Range[startCell, endCell];
            }
            finally
            {
                ReleaseComObject(endCell);
                ReleaseComObject(startCell);
            }
        }

        private void SetRangeValue(Worksheet worksheet, int column, object[,] values)
        {
            Range range = null;

            try
            {
                range = GetColumnRange(worksheet, column);
                range.Value2 = values;
            }
            finally
            {
                ReleaseComObject(range);
            }
        }

        private void SetRangeFormula(Worksheet worksheet, int column, object[,] formulas)
        {
            Range range = null;

            try
            {
                range = GetColumnRange(worksheet, column);
                range.Formula = formulas;
            }
            finally
            {
                ReleaseComObject(range);
            }
        }

        private static void SetCellValue(Worksheet worksheet, int row, int column, object value)
        {
            Range cell = null;

            try
            {
                cell = (Range)worksheet.Cells[row, column];
                cell.Value2 = value;
            }
            finally
            {
                ReleaseComObject(cell);
            }
        }

        private static void SetCellFormula(Worksheet worksheet, int row, int column, string formula)
        {
            Range cell = null;

            try
            {
                cell = (Range)worksheet.Cells[row, column];
                cell.Formula = formula;
            }
            finally
            {
                ReleaseComObject(cell);
            }
        }

        private static void SetCellFormulaLocal(Worksheet worksheet, int row, int column, string formula)
        {
            Range cell = null;

            try
            {
                cell = (Range)worksheet.Cells[row, column];
                cell.FormulaLocal = formula;
            }
            finally
            {
                ReleaseComObject(cell);
            }
        }

        private static void SetCellNumberFormat(Worksheet worksheet, int row, int column, string numberFormat)
        {
            Range cell = null;

            try
            {
                cell = (Range)worksheet.Cells[row, column];
                cell.NumberFormat = numberFormat;
            }
            finally
            {
                ReleaseComObject(cell);
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
    }
}
