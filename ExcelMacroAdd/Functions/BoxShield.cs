﻿using ExcelMacroAdd.Servises;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal class BoxShield : AbstractFunctions
    {
        private readonly Lazy<DBConect> dBConect;
        public BoxShield(Lazy<DBConect> dBConect)
        {
            this.dBConect = dBConect;
        }

        public override void Start()
        {
            int firstRow, countRow, endRow;
            try
            {
                // Открываем соединение с базой данных    
                dBConect.Value.OpenDB();

                if (application.ActiveWorkbook.Name == dBConect.Value.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2))            // Проверка по имени книги
                {
                    firstRow = cell.Row;                 // Вычисляем верхний элемент
                    countRow = cell.Rows.Count;          // Вычисляем кол-во выделенных строк
                    endRow = firstRow + countRow;
                    // Инициализируем структуру для записи                 
                    do
                    {
                        string sArticle = Convert.ToString(worksheet.Cells[firstRow, 26].Value2);
                        string query = "SELECT * FROM base WHERE article = '" + sArticle + "'";

                        if (dBConect.Value.CheckReadDB(query))
                        {
                            worksheet.get_Range("Z" + firstRow).Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
                        }
                        else
                        {
                            // Передеем структуру по референсной ссылке в библиотечный метод 
                            var table = dBConect.Value.ReadingDB(query);
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
                }
                else
                {
                    MessageBox.Show(
                    "Программа работает только в файле " + dBConect.Value.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2) + "\n Пожайлуста откройте целевую книгу и запустите программу.",
                    "Ошибка вызова",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                }
                // Закрываем соединение с базой данных
                dBConect.Value.CloseDB();
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
