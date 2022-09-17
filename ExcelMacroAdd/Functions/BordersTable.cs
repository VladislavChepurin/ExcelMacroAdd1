using Microsoft.Office.Interop.Excel;

namespace ExcelMacroAdd.Functions
{
    internal class BordersTable : AbstractFunctions
    {
        public sealed override void Start()
        {
            var excelcells = application.Selection;
            XlBordersIndex borderIndex;

            borderIndex = XlBordersIndex.xlEdgeLeft; //Левая граница
            excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelcells.Borders[borderIndex].ColorIndex = 0;

            borderIndex = XlBordersIndex.xlEdgeTop; //Верхняя граница
            excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelcells.Borders[borderIndex].ColorIndex = 0;

            borderIndex = XlBordersIndex.xlEdgeBottom; //Нижняя граница
            excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelcells.Borders[borderIndex].ColorIndex = 0;

            borderIndex = XlBordersIndex.xlEdgeRight;  //Правая граница
            excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelcells.Borders[borderIndex].ColorIndex = 0;

            borderIndex = XlBordersIndex.xlInsideHorizontal;  //Внутренняя горизонтальня граница
            excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelcells.Borders[borderIndex].ColorIndex = 0;

            borderIndex = XlBordersIndex.xlInsideVertical;  //Внутренняя горизонтальня граница
            excelcells.Borders[borderIndex].Weight = XlBorderWeight.xlThin;
            excelcells.Borders[borderIndex].LineStyle = XlLineStyle.xlContinuous;
            excelcells.Borders[borderIndex].ColorIndex = 0;
        }
    }
}
