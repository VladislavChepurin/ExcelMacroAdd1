using ExcelMacroAdd.Serializable.Entity.Interfaces;
using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class TypeNky : UserControl
    {
        private readonly ITypeNkySettings[] typeNkySettings;
        private readonly DataGridView nkyDataGridView = new DataGridView();       

        public TypeNky(ITypeNkySettings[] typeNkySettings)
        {
            this.typeNkySettings = typeNkySettings;
            InitializeComponent();
        }

        private void TypeNky_Load(object sender, EventArgs e)
        {
            SetupDataGridView();
            SetupLayout();
            PopulateDataGridView();
        }

        private void PopulateDataGridView()
        {
            foreach (var item in typeNkySettings)
            {
                string[] row = { item.Number.ToString(), item.Description, item.BuildTime.ToString() };
                nkyDataGridView.Rows.Add(row);
            }
        }     

        private void SetupLayout()
        {
            this.Size = new Size(600, 500);       
        }

        private void SetupDataGridView()
        {
            this.Controls.Add(nkyDataGridView);
            nkyDataGridView.ColumnCount = 3;

            nkyDataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            nkyDataGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            nkyDataGridView.ColumnHeadersDefaultCellStyle.Font =
                new System.Drawing.Font(nkyDataGridView.Font, FontStyle.Bold);

            nkyDataGridView.Name = "nkyDataGridView";
            nkyDataGridView.Location = new System.Drawing.Point(5, 5);
            nkyDataGridView.Size = new Size(380, 500);

            nkyDataGridView.AutoSizeRowsMode =
                DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            nkyDataGridView.ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Single;
            nkyDataGridView.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            nkyDataGridView.GridColor = Color.Black;
            nkyDataGridView.RowHeadersVisible = false;

            nkyDataGridView.Columns[0].Name = "№";
            nkyDataGridView.Columns[1].Name = "Тип изделия";
            nkyDataGridView.Columns[2].Name = "Норматив";

            nkyDataGridView.Columns[0].Width = 25;
            nkyDataGridView.Columns[1].Width = 290;
            nkyDataGridView.Columns[2].Width = 65;

            nkyDataGridView.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;
            nkyDataGridView.MultiSelect = false;
            nkyDataGridView.Dock = DockStyle.Fill;

            nkyDataGridView.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f => f.SortMode = DataGridViewColumnSortMode.NotSortable);
            nkyDataGridView.ReadOnly = true;

            nkyDataGridView.BackgroundColor = Color.White;
        }
    }
}
