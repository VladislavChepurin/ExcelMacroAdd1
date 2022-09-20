using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.Interfaces;
using System;
using System.Data;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal class BoxShield : AbstractFunctions
    {
        private readonly IResources resources;
        private readonly IJornalData jornalData;

        public BoxShield(IJornalData jornalData ,IResources resources)
        {
            this.jornalData = jornalData;
            this.resources = resources;
        }

        public sealed override async void Start()
        {
            if (application.ActiveWorkbook.Name != resources.NameFileJornal) // Проверка по имени книги
            {
                MessageWarning("Функция работает только в \"Журнале учета НКУ\" текущего года. \n Пожайлуста откройте необходимую книгу Excel.",
                    "Имя книги не совпадает с целевой");
                return;
            }
            int firstRow, countRow, endRow;
            firstRow = cell.Row;                 // Вычисляем верхний элемент
            countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
            endRow = firstRow + countRow;
            do
            {
                try
                {
                    string sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);
                    var jornalNKU = await jornalData.GetEntityJornal(sArticle);

                    if (jornalNKU is null)
                    {
                        worksheet.get_Range("Z" + firstRow).Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
                        firstRow++;
                        continue;
                    }
                    worksheet.get_Range("K" + firstRow).Value2 = jornalNKU.Ip.ToString() ?? String.Empty;
                    worksheet.get_Range("L" + firstRow).Value2 = jornalNKU.Klima ?? String.Empty;
                    worksheet.get_Range("M" + firstRow).Value2 = jornalNKU.Reserve ?? String.Empty;
                    worksheet.get_Range("N" + firstRow).Value2 = jornalNKU.Height ?? String.Empty;
                    worksheet.get_Range("O" + firstRow).Value2 = jornalNKU.Width ?? String.Empty;
                    worksheet.get_Range("P" + firstRow).Value2 = jornalNKU.Depth ?? String.Empty;
                    worksheet.get_Range("AC" + firstRow).Value2 = jornalNKU.Execution ?? String.Empty;
                }
                catch (DataException)
                {
                    MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                        "Ошибка базы данных");
                    return;
                }
                catch (Exception e)
                {
                    MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {e.Message}",
                        "Ошибка базы данных");
                    return;
                }
                firstRow++;
            }
            while (endRow > firstRow);
        }
    }
}
