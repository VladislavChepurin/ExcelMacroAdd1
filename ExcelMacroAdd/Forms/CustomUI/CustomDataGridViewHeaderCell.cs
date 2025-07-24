using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms.CustomUI
{
    public class CustomDataGridViewHeaderCell : DataGridViewColumnHeaderCell
    {
        private SortOrder _sortGlyphDirection = SortOrder.None;
        private readonly string _headerText = string.Empty;

        public CustomDataGridViewHeaderCell(string headerText)
        {
            _headerText = headerText;
        }

        [Browsable(false)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public new SortOrder SortGlyphDirection
        {
            get => _sortGlyphDirection;
            set
            {
                if (_sortGlyphDirection != value)
                {
                    _sortGlyphDirection = value;
                    // Принудительно перерисовываем ячейку при изменении значения
                    DataGridView?.InvalidateCell(this);
                }
            }
        }

        protected override void Paint(
            Graphics graphics,
            Rectangle clipBounds,
            Rectangle cellBounds,
            int rowIndex,
            DataGridViewElementStates dataGridViewElementState,
            object value,
            object formattedValue,
            string errorText,
            DataGridViewCellStyle cellStyle,
            DataGridViewAdvancedBorderStyle advancedBorderStyle,
            DataGridViewPaintParts paintParts)
        {
            // Рисуем стандартные элементы (фон, границы)
            base.Paint(graphics, clipBounds, cellBounds, rowIndex, dataGridViewElementState,
                value, formattedValue, errorText, cellStyle, advancedBorderStyle,
                paintParts & ~DataGridViewPaintParts.ContentForeground);

            
            // Рисуем текст заголовка
            if (!string.IsNullOrEmpty(_headerText))
            {
                DrawHeaderText(graphics, cellBounds, _headerText, cellStyle);
            }

            // Рисуем значок сортировки
            if (_sortGlyphDirection != SortOrder.None)
            {
                DrawSortIcon(graphics, cellBounds, _sortGlyphDirection);
            }
        }

        private void DrawHeaderText(Graphics graphics, Rectangle cellBounds, string text, DataGridViewCellStyle cellStyle)
        {              
            // Определяем область для текста с учетом отступов
            Rectangle textBounds = new Rectangle(
                cellBounds.Left + 5,
                cellBounds.Top,
                cellBounds.Width - 20, // Оставляем место для значка сортировки
                cellBounds.Height);

            Brush textBrush = Brushes.Black;

            graphics.DrawString(
            text,
            cellStyle.Font,
            textBrush,
            textBounds,
            new StringFormat
            {
                Alignment = StringAlignment.Near,
                LineAlignment = StringAlignment.Center,
                Trimming = StringTrimming.EllipsisCharacter,
                FormatFlags = StringFormatFlags.NoWrap
            }
            );
        }

        private void DrawSortIcon(Graphics graphics, Rectangle cellBounds, SortOrder direction)
        {
            int iconSize = 6;
            int padding = 5;
            int x = cellBounds.Right - iconSize - padding;
            int y = cellBounds.Top + (cellBounds.Height - iconSize) / 2;

            if (direction == SortOrder.Ascending)
            {
                // Стрелка вверх
                Point[] points = {
                new Point(x, y + iconSize),
                new Point(x + iconSize / 2, y),
                new Point(x + iconSize, y + iconSize)
            };
                graphics.FillPolygon(Brushes.Brown, points);
            }
            else
            {
                // Стрелка вниз
                Point[] points = {
                new Point(x, y),
                new Point(x + iconSize / 2, y + iconSize),
                new Point(x + iconSize, y)
            };
                graphics.FillPolygon(Brushes.Brown, points);
            }
        }
    }
}
