using ExcelMacroAdd.Serializable.Entity.Interfaces;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;


namespace ExcelMacroAdd.Functions
{
    internal sealed class Linker : AbstractFunctions
    {
        private readonly ICorrectFontResources correctFontResources;

        public Linker(ICorrectFontResources correctFontResources)
        {
            this.correctFontResources = correctFontResources;
        }

        public override void Start()
        {
            //Создаем коллекцию
            List<int> list = new List<int>();
            // В коллекцию помещаем парсированые имена листов
            foreach (Worksheet sheet in WorkBook.Sheets)
            {
                int.TryParse(sheet.Name, out int result);
                list.Add(result);
            }
            //Проверяем на налиечие этих номеров в коллекции, если есть, то имя вкладке не присваиваем                 
            if (!list.Contains(Worksheet.Index - 1))
            {
                //имя вкладки в зависимости от индекса
                Worksheet.Name = (Worksheet.Index - 1).ToString();
            }

            // Заголовки столбцов
            var headers = new List<(string Cell, string Value)>
                {
                    ("A1", "Артикул"),
                    ("B1", "Описание"),
                    ("C1", "Кол-во"),
                    ("D1", "Кратность"),
                    ("E1", "Пр-ль"),
                    ("F1", "Скидка"),
                    ("G1", "Цена"),
                    ("H1", "Цена со скидкой"),
                    ("I1", "Стоимость"),
                    ("J1", "Дата и время")
                };

            foreach (var header in headers)
            {
                Worksheet.Range[header.Cell].Value2 = header.Value;
            }

            // Ширина столбцов
            var columnWidths = new Dictionary<string, double>
            {
                ["A:A"] = 21,
                ["B:B"] = 80,
                ["C:C"] = 10,
                ["D:I"] = 13,
                ["J:J"] = 0
            };

            foreach (var width in columnWidths)
            {
                Worksheet.Range[width.Key].ColumnWidth = width.Value;
            }         

            //размечаем границы и правим шрифты
            Worksheet.Range["A1", "J500"].Cells.Font.Name = correctFontResources.NameFont;
            Worksheet.Range["A1", "J500"].Cells.Font.Size = correctFontResources.SizeFont;

            // Форматирование
            var fontRange = Worksheet.Range["A1", "J100"];  
            fontRange.Font.Name = correctFontResources.NameFont;
            fontRange.Font.Size = correctFontResources.SizeFont;

            var bordersRange = Worksheet.Range["A1", "J11"];
            bordersRange.Borders.LineStyle = XlLineStyle.xlContinuous;  // Добавлено оформление границ        
        }
    }
}
