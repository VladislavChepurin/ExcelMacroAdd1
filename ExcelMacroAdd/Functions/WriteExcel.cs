using ExcelMacroAdd.Services;
using ExcelMacroAdd.Services.Interfaces;
using System;
using System.Runtime.InteropServices;

//Rewiew OK 22.04.2025
namespace ExcelMacroAdd.Functions
{
    internal sealed class WriteExcel : AbstractFunctions
    {
        private readonly IDataInXml _dataInXml;
        private readonly string _vendor;
        private readonly int _rowsOffset;
        private readonly string _article;
        private readonly int _amount;

        private const int ArticleColumn = 1;
        private const int DescriptionColumn = 2;
        new private const int  QuantityColumn = 3;
        private const int MultiplicityColumn = 4;
        private const int VendorColumn = 5;
        private const int DiscountColumn = 6;
        private const int PriceColumn = 7;
        private const int TotalPriceColumn = 8;
        private const int CoastColumn = 9;
        private const int DateColumn = 10;

        public WriteExcel(IDataInXml dataInXml, string vendor, int rowsOffset = 0, string article = null, int amount = 0)
        {
            _dataInXml = dataInXml ?? throw new ArgumentNullException(nameof(dataInXml));
            _vendor = vendor ?? throw new ArgumentNullException(nameof(vendor));
            _rowsOffset = rowsOffset;
            _article = article;
            _amount = amount;
        }

        public override void Start()
        {
            try
            {
                if (Worksheet == null || Cell == null)
                    throw new InvalidOperationException("Не инициализирован объект Excel.");

                var vendors = _dataInXml.ReadFileXml();
                var vendorData = _dataInXml.ReadElementXml(_vendor, vendors)
                    ?? throw new ArgumentException($"Вендор {_vendor} не найден.");

                int startRow = Cell.Row;
                int totalRows = (_article != null) ? 1 : Cell.Rows.Count;

                var data = new object[totalRows, 10]; // 10 столбцов

                for (int i = 0; i < totalRows; i++)
                {
                    int currentRow = startRow + _rowsOffset + i;

                    Worksheet.Cells[currentRow, ArticleColumn] = _article;
                    Worksheet.Cells[currentRow, DescriptionColumn].FormulaLocal = string.Format(vendorData.Formula_1, currentRow);
                    Worksheet.Cells[currentRow, QuantityColumn] = _amount != 0 ? (int?)_amount : null;
                    Worksheet.Cells[currentRow, MultiplicityColumn].FormulaLocal = string.Format(vendorData.Formula_2, currentRow);
                    Worksheet.Cells[currentRow, VendorColumn] = _vendor;
                    Worksheet.Cells[currentRow, DiscountColumn] = vendorData.Discount;
                    Worksheet.Cells[currentRow, PriceColumn].FormulaLocal = string.Format(vendorData.Formula_3, currentRow);
                    Worksheet.Cells[currentRow, TotalPriceColumn].Formula = $"=G{currentRow}*(100-F{currentRow})/100";
                    Worksheet.Cells[currentRow, CoastColumn].Formula = $"=H{currentRow}*C{currentRow}";
                    Worksheet.Cells[currentRow, DateColumn].NumberFormat = "ДД.ММ.ГГ ч:мм";
                    Worksheet.Cells[currentRow, DateColumn] = DateTime.Now;                   
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
                Marshal.FinalReleaseComObject(Worksheet);
            }
        }
    }
}
