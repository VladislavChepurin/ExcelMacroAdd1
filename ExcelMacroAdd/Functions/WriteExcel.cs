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
            _startRow = Cell.Row;
            _countRows = Cell.Rows.Count;
        }

        public WriteExcel(IDataInXml dataInXml, string vendor, string article, int startOffset = 0, int amount = 0)
        {
            _dataInXml = dataInXml ?? throw new ArgumentNullException(nameof(dataInXml));
            _vendor = vendor ?? throw new ArgumentNullException(nameof(vendor));
            _startRow = Cell.Row + startOffset;
            _countRows = 1;
            _article = article;
            _amount = amount;
        }

        public override void Start()
        {
            if (ExcelPerformanceScope.CurrentNestingLevel == 0)
                throw new InvalidOperationException(
                    "WriteExcel.Start() must be called inside ExcelPerformanceScope");
            try
            {
                if (Worksheet == null || Cell == null)
                    throw new InvalidOperationException("Не инициализирован объект Excel.");

                var vendors = _dataInXml.ReadFileXml();
                var vendorData = _dataInXml.ReadElementXml(_vendor, vendors)
                    ?? throw new ArgumentException($"Вендор {_vendor} не найден.");

                if (_countRows <= 1)
                {
                    WriteSingleRow(_startRow, vendorData);
                }
                else
                {
                    WriteBatchRows(vendorData);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Запись одной строки.
        /// Порядок: сначала все лёгкие значения, формулы ВПР — в самом конце.
        /// </summary>
        private void WriteSingleRow(int row, UserVariables.Vendor vendorData)
        {
            // 1) Лёгкие значения — мгновенно
            if (_article != null)
                Worksheet.Cells[row, ArticleColumn] = _article;
            if (_amount != 0)
                Worksheet.Cells[row, QuantityColumn] = _amount;

            Worksheet.Cells[row, VendorColumn] = _vendor;
            Worksheet.Cells[row, DiscountColumn] = vendorData.Discount;
            Worksheet.Cells[row, DateColumn].NumberFormat = "ДД.ММ.ГГ ч:мм";
            Worksheet.Cells[row, DateColumn] = DateTime.Now;

            // 2) Простые формулы (не ВПР) — быстрые
            Worksheet.Cells[row, TotalPriceColumn].Formula = $"=G{row}*(100-F{row})/100";
            Worksheet.Cells[row, CoastColumn].Formula = $"=H{row}*C{row}";

            // 3) Формулы ВПР — тяжёлые, в конце
            Worksheet.Cells[row, DescriptionColumn].FormulaLocal = string.Format(vendorData.Formula_1, row);
            Worksheet.Cells[row, MultiplicityColumn].FormulaLocal = string.Format(vendorData.Formula_2, row);
            Worksheet.Cells[row, PriceColumn].FormulaLocal = string.Format(vendorData.Formula_3, row);
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
        private void WriteBatchRows(UserVariables.Vendor vendorData)
        {
            int endRow = _startRow + _countRows - 1;
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

            SetRangeValue(VendorColumn, vendorValues);
            SetRangeValue(DiscountColumn, discountValues);

            Range dateRange = GetColumnRange(DateColumn);
            dateRange.NumberFormat = "ДД.ММ.ГГ ч:мм";
            dateRange.Value2 = dateValues;
            Marshal.ReleaseComObject(dateRange);

            // === Простые формулы (не ВПР) ===
            var totalPriceFormulas = new object[_countRows, 1];
            var coastFormulas = new object[_countRows, 1];

            for (int i = 0; i < _countRows; i++)
            {
                int row = _startRow + i;
                totalPriceFormulas[i, 0] = $"=G{row}*(100-F{row})/100";
                coastFormulas[i, 0] = $"=H{row}*C{row}";
            }

            SetRangeFormula(TotalPriceColumn, totalPriceFormulas);
            SetRangeFormula(CoastColumn, coastFormulas);

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
            WriteAsText(DescriptionColumn, formula1Text);
            WriteAsText(MultiplicityColumn, formula2Text);
            WriteAsText(PriceColumn, formula3Text);

            // Фаза 2: конвертируем текст → формулы (одна операция на столбец)
            ActivateFormulas(DescriptionColumn);
            ActivateFormulas(MultiplicityColumn);
            ActivateFormulas(PriceColumn);
        }

        /// <summary>
        /// Записывает строки-формулы как текст. 
        /// NumberFormat = "@" заставляет Excel трактовать "=ВПР(...)" как литерал.
        /// Value2 со строками в текстовом формате — мгновенная операция.
        /// </summary>
        private void WriteAsText(int column, object[,] formulaTexts)
        {
            Range range = GetColumnRange(column);
            range.NumberFormat = "@";
            range.Value2 = formulaTexts;
            Marshal.ReleaseComObject(range);
        }

        /// <summary>
        /// Конвертирует текстовые строки формул в реальные формулы Excel.
        /// Снимает текстовый формат, читает значения и записывает обратно
        /// как FormulaLocal — одна COM-операция на весь диапазон.
        /// </summary>
        private void ActivateFormulas(int column)
        {
            Range range = GetColumnRange(column);
            try
            {
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
                Marshal.ReleaseComObject(range);
            }
        }

        // === Вспомогательные методы ===

        private Range GetColumnRange(int column)
        {
            int endRow = _startRow + _countRows - 1;
            return Worksheet.Range[
                Worksheet.Cells[_startRow, column],
                Worksheet.Cells[endRow, column]];
        }

        private void SetRangeValue(int column, object[,] values)
        {
            Range range = GetColumnRange(column);
            range.Value2 = values;
            Marshal.ReleaseComObject(range);
        }

        private void SetRangeFormula(int column, object[,] formulas)
        {
            Range range = GetColumnRange(column);
            range.Formula = formulas;
            Marshal.ReleaseComObject(range);
        }
    }
}