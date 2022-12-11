using Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal sealed class BordersTable : AbstractFunctions
    {
        public override void Start()
        {
            var excelCells = Application.Selection;

            var borderIndex = XlBordersIndex.xlEdgeLeft; //Левая граница
            excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelCells.Borders[borderIndex].ColorIndex = 0;

            borderIndex = XlBordersIndex.xlEdgeTop; //Верхняя граница
            excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelCells.Borders[borderIndex].ColorIndex = 0;

            borderIndex = XlBordersIndex.xlEdgeBottom; //Нижняя граница
            excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelCells.Borders[borderIndex].ColorIndex = 0;

            borderIndex = XlBordersIndex.xlEdgeRight;  //Правая граница
            excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelCells.Borders[borderIndex].ColorIndex = 0;

            borderIndex = XlBordersIndex.xlInsideHorizontal;  //Внутренняя горизонтальня граница
            excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelCells.Borders[borderIndex].ColorIndex = 0;

            borderIndex = XlBordersIndex.xlInsideVertical;  //Внутренняя горизонтальня граница
            excelCells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelCells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelCells.Borders[borderIndex].ColorIndex = 0;
        }
    }
}
