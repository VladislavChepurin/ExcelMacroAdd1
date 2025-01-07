using ExcelMacroAdd.Forms.SupportiveFunction;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ButtonForm = System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class TypeNky : UserControl
    {
        private readonly ITypeNkySettings[] typeNkySettings;
        private readonly Panel buttonPanel = new Panel();
        private readonly DataGridView nkyDataGridView = new DataGridView();
        private readonly ButtonForm.Button addTypeButton = new ButtonForm.Button();
        private readonly ButtonForm.Button deleteTypeButton = new ButtonForm.Button();

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

        private void addTypeButton_Click(object sender, EventArgs e)
        {
            if (nkyDataGridView.SelectedCells.Count > 0)
            {
                int selectedrowindex = nkyDataGridView.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = nkyDataGridView.Rows[selectedrowindex];
                string cellValue = selectedRow.Cells[0].Value.ToString();

                var typeNky = new AddTypeNky(cellValue);
                typeNky.Start();
            }
        }

        private void deleteTypeButton_Click(object sender, EventArgs e)
        {
            var typeNky = new DeleteTypeNky();
            typeNky.Start();
        }

        private void SetupLayout()
        {
            this.Size = new Size(600, 500);

            addTypeButton.Text = "Добавить тип";
            addTypeButton.Location = new System.Drawing.Point(50, 40);
            addTypeButton.Width = 90;
            addTypeButton.Height = 30;
            addTypeButton.Click += new EventHandler(addTypeButton_Click);

            deleteTypeButton.Text = "Удалить тип";
            deleteTypeButton.Location = new System.Drawing.Point(245, 40);
            deleteTypeButton.Width = 90;
            deleteTypeButton.Height = 30;
            deleteTypeButton.Click += new EventHandler(deleteTypeButton_Click);

            buttonPanel.Controls.Add(addTypeButton);
            buttonPanel.Controls.Add(deleteTypeButton);
            buttonPanel.Height = 100;
            buttonPanel.Dock = DockStyle.Bottom;

            this.Controls.Add(this.buttonPanel);
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
