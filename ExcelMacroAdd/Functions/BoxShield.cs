using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.Servises;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal class BoxShield : AbstractFunctions
    {
        private readonly IDBConect dBConect;
        public BoxShield(IDBConect dBConect)
        {
            this.dBConect = dBConect;
        }

        public override void Start()
        {
            dBConect?.OpenDB();
            if (application.ActiveWorkbook.Name != dBConect?.ReadOnlyOneNoteDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2)) // Проверка по имени книги
            {
                MessageWrongNameJournal();
                dBConect?.CloseDB();
                return;
            }

            int firstRow, countRow, endRow;
            try
            {
                firstRow = cell.Row;                 // Вычисляем верхний элемент
                countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
                endRow = firstRow + countRow;
                // Инициализируем структуру для записи                 
                do
                {
                    string sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);
                    string query = $"SELECT * FROM base WHERE article = '{sArticle}';";
                    //Если не возвращается значение, то этой записи нет
                    //Костыль но работает, прекрасно
                    //Точно ли необходим этот метод? Или можно обойтись одним ReadSeveralNotesDB.
                    if (dBConect?.ReadOnlyOneNoteDB(query, 1) is null)
                    {
                        worksheet.get_Range("Z" + firstRow).Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
                    }
                    else
                    {
                        // Передеем структуру по референсной ссылке в библиотечный метод 
                        var table = dBConect.ReadSeveralNotesDB(query);
                        // Присваеваем ячейкам данные из массива
                        worksheet.get_Range("K" + firstRow).Value2 = table.IpTable ?? String.Empty;
                        worksheet.get_Range("L" + firstRow).Value2 = table.KlimaTable ?? String.Empty;
                        worksheet.get_Range("M" + firstRow).Value2 = table.ReserveTable ?? String.Empty;
                        worksheet.get_Range("N" + firstRow).Value2 = table.HeightTable ?? String.Empty;
                        worksheet.get_Range("O" + firstRow).Value2 = table.WidthTable ?? String.Empty;
                        worksheet.get_Range("P" + firstRow).Value2 = table.DepthTable ?? String.Empty;
                        worksheet.get_Range("AC" + firstRow).Value2 = table.ExecutionTable ?? String.Empty;
                    }
                    firstRow++;
                }
                while (endRow > firstRow);
                // Закрываем соединение с базой данных
                dBConect?.CloseDB();
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                exception.ToString(),
                "Ошибка надстройки",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
            }
        }
    }
}
