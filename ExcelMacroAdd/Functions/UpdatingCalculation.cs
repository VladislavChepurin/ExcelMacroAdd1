using ExcelMacroAdd.Models;
using ExcelMacroAdd.Models.Interface;
using ExcelMacroAdd.Services;
using ExcelMacroAdd.Services.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;

namespace ExcelMacroAdd.Functions
{
    internal class UpdatingCalculation : AbstractFunctions
    {
        private readonly IDataInXml dataInXml;

        public UpdatingCalculation(IDataInXml dataInXml)
        {
            this.dataInXml = dataInXml;
        }

        public override void Start()
        {
            Worksheet.Range["A1", Type.Missing].Select();   //Фокус на ячейку А1, костыль
            int currentRow = 2;
            int counter = 0;
            string currentVendor;

            while (true)
            {
                if (Worksheet.Range["A" + currentRow].Value != null || Worksheet.Range["B" + currentRow].Value != null)
                {
                    currentVendor = Convert.ToString(Worksheet.Cells[currentRow, 5].Value2);
                    UpdateCalc(new Iek(), currentVendor, currentRow);
                    UpdateCalc(new Ekf(), currentVendor, currentRow);
                    UpdateCalc(new Dkc(), currentVendor, currentRow);
                    UpdateCalc(new Keaz(), currentVendor, currentRow);
                    UpdateCalc(new Dekraft(), currentVendor, currentRow);
                    //UpdateCalc(new Tdm(), currentVendor, currentRow);
                    //UpdateCalc(new Abb(), currentVendor, currentRow);
                    //UpdateCalc(new Schneider(), currentVendor, currentRow);
                    UpdateCalc(new Chint(), currentVendor, currentRow);
                    counter = 0;
                }
                else if (counter > 1)
                {
                    break;
                }
                else
                {
                    counter++;
                }
                currentRow++;           
            }               
        }

        private void UpdateCalc(Vendors vendors, string currentVendor, int rowsLine)
        {
            if (vendors.RangeSearch.Contains(currentVendor, new CustomStringComparer())) {
                WriteExcel writeExcel = new WriteExcel(dataInXml, vendors.OutValue, rowsLine - 2);
                writeExcel.Start();
            }
        }
    }
}
