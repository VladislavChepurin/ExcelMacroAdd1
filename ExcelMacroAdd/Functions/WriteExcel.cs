using ExcelMacroAdd.Services;
using ExcelMacroAdd.Services.Interfaces;
using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.InteropServices;

//Rewiew OK 22.04.2025
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
        new private const int  QuantityColumn = 3;
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
            try
            {
                if (Worksheet == null || Cell == null)
                    throw new InvalidOperationException("Не инициализирован объект Excel.");

                var vendors = _dataInXml.ReadFileXml();
                var vendorData = _dataInXml.ReadElementXml(_vendor, vendors)
                    ?? throw new ArgumentException($"Вендор {_vendor} не найден.");

                int currentRow;

                for (int i = 0; i < _countRows; i++)
                {
                    currentRow = _startRow + i;
                    if (_article != null)
                        Worksheet.Cells[currentRow, ArticleColumn] = _article;
                    Worksheet.Cells[currentRow, DescriptionColumn].FormulaLocal = string.Format(vendorData.Formula_1, currentRow);
                    if (_amount != 0)
                        Worksheet.Cells[currentRow, QuantityColumn] = _amount;
                    Worksheet.Cells[currentRow, MultiplicityColumn].FormulaLocal = string.Format(vendorData.Formula_2, currentRow);
                    Worksheet.Cells[currentRow, VendorColumn] = _vendor;
                    Worksheet.Cells[currentRow, DiscountColumn] = vendorData.Discount;
                    Worksheet.Cells[currentRow, PriceColumn].FormulaLocal = string.Format(vendorData.Formula_3, _startRow);
                    Worksheet.Cells[currentRow, TotalPriceColumn].Formula = $"=G{_startRow}*(100-F{currentRow})/100";
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
