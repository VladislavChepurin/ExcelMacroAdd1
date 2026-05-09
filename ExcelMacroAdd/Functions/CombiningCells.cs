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
            Range selectedRange = null;
            Range activeCell = null;

            try
            {
                selectedRange = Cell;
                var cellValue = selectedRange.Value;

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

                    // Очищаем и обновляем значение
                    selectedRange.ClearContents();
                    activeCell = Application.ActiveCell as Range;
                    if (activeCell != null)
                    {
                        activeCell.Value2 = string.Join(";\n", elements);
                    }
                }
            }
            finally
            {
                ReleaseComObject(activeCell);
                ReleaseComObject(selectedRange);
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
