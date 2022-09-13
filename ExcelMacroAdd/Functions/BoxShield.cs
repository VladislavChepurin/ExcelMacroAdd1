using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Interfaces;
using System;
using System.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal class BoxShield : AbstractFunctions
    {
        private JornalNKU jornalNKU;
        private readonly IResources resources;

        public BoxShield(IResources resources)
        {
            this.resources = resources;
        }

        protected internal override void Start()
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
            using (DataContext db = new DataContext())
            {
                var jornalNKUs = db.JornalNKU;
                do
                {                    
                    try
                    {
                        string sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);
                        jornalNKU = jornalNKUs.Where(p => p.Article == sArticle).FirstOrDefault();
                        if (jornalNKU is null)
                        {
                            worksheet.get_Range("Z" + firstRow).Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
                            firstRow++;
                            continue;
                        }
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

                    worksheet.get_Range("K" + firstRow).Value2 = jornalNKU.Ip.ToString() ?? String.Empty;
                    worksheet.get_Range("L" + firstRow).Value2 = jornalNKU.Klima ?? String.Empty;
                    worksheet.get_Range("M" + firstRow).Value2 = jornalNKU.Reserve ?? String.Empty;
                    worksheet.get_Range("N" + firstRow).Value2 = jornalNKU.Height ?? String.Empty;
                    worksheet.get_Range("O" + firstRow).Value2 = jornalNKU.Width ?? String.Empty;
                    worksheet.get_Range("P" + firstRow).Value2 = jornalNKU.Depth ?? String.Empty;
                    worksheet.get_Range("AC" + firstRow).Value2 = jornalNKU.Execution ?? String.Empty;

                    firstRow++;
                }
                while (endRow > firstRow);
            }
        }
    }
}
