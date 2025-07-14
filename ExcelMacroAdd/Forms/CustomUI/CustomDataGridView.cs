using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms.CustomUI
{
    class CustomDataGridView : DataGridView
    {

        public CustomDataGridView() : base()
        {
            VerticalScrollBar.Visible = true;
            VerticalScrollBar.VisibleChanged += new EventHandler(VerticalScrollBar_VisibleChanged);
        }

        void VerticalScrollBar_VisibleChanged(object sender, EventArgs e)
        {
            if (!VerticalScrollBar.Visible)
            {
                VerticalScrollBar.Location = new Point(ClientRectangle.Width - VerticalScrollBar.Width - 1, 1);
                VerticalScrollBar.Height = this.ClientRectangle.Height - 2;
                VerticalScrollBar.Width = VerticalScrollBar.Width;
                VerticalScrollBar.Show();
            }
        }
    }
}
