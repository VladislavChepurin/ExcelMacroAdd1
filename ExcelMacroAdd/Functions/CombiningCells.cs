using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace ExcelMacroAdd.Functions
{
    internal class CombiningCells : AbstractFunctions

    {
        public override void Start()
        {
            var firstRow = Cell.Row; // Вычисляем верхний элемент
            var collumn = Cell.Column;
            var countRow = Cell.Rows.Count; // Вычисляем кол-во выделенных строк                                 

            Worksheet.Cells[firstRow, collumn].Value2 =
                String.Join(" ", Main(firstRow, countRow, collumn).ToArray());                                                          
        }

        private List<string> Main(int firstRow, int countRow, int collumn)
        {
            var endRow = firstRow + countRow;
            List<string> list = new List<string>();
           
            if (countRow > 1)
            {
                // Цикл переборки строк
                do
                {
                    string value = Worksheet.Cells[firstRow, collumn].Text as string;

                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        list.Add(value);
                        Worksheet.Cells[firstRow, collumn].Value2 = null;
                    }

                    firstRow++;
                }
                while (endRow > firstRow);
            }
            return list;
        }
    }
}
