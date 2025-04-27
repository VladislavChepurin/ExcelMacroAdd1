using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ExcelMacroAdd.Functions
{
    internal class CombiningCells : AbstractFunctions
    {
        public override void Start()
        {
            // Получаем значение ячейки как двумерный массив
            var cellValue = Cell.Value;
            if (cellValue is Object[,] array2D)
            {
                // Создаем список для элементов с обработкой null
                var elements = new List<string>(array2D.Length);

                // Обрабатываем все элементы массива
                foreach (var item in array2D)
                {
                    if (item is null) continue;
                    elements.Add(item.ToString());
                }

                // Получаем исходный диапазон ячейки
                var targetRange = Worksheet.Range[Cell.Address];

                try
                {
                    // Очищаем и обновляем значение
                    targetRange.ClearContents();
                    Application.ActiveCell.Value2 = string.Join(";\n", elements);
                }
                finally
                {
                    // Освобождаем COM-объекты
                    Marshal.ReleaseComObject(targetRange);
                }
            }
        }
    }
}
