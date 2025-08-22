using ExcelMacroAdd.Forms.CustomUI;

namespace ExcelMacroAdd.Forms
{
    partial class NotPriceComponents
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnWritingToSheet = new System.Windows.Forms.Button();
            this.btnAddRecord = new System.Windows.Forms.Button();
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.btnDeleteRecord = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnUpdateRecord = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.dataGridView = new ExcelMacroAdd.Forms.CustomUI.CustomDataGridView();
            this.linkToTheWebsite = new System.Windows.Forms.LinkLabel();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // btnWritingToSheet
            // 
            this.btnWritingToSheet.Location = new System.Drawing.Point(10, 520);
            this.btnWritingToSheet.Name = "btnWritingToSheet";
            this.btnWritingToSheet.Size = new System.Drawing.Size(75, 25);
            this.btnWritingToSheet.TabIndex = 0;
            this.btnWritingToSheet.Text = "На лист";
            this.btnWritingToSheet.UseVisualStyleBackColor = true;
            // 
            // btnAddRecord
            // 
            this.btnAddRecord.Location = new System.Drawing.Point(626, 520);
            this.btnAddRecord.Name = "btnAddRecord";
            this.btnAddRecord.Size = new System.Drawing.Size(75, 25);
            this.btnAddRecord.TabIndex = 1;
            this.btnAddRecord.Text = "Добавить";
            this.btnAddRecord.UseVisualStyleBackColor = true;
            // 
            // searchTextBox
            // 
            this.searchTextBox.Location = new System.Drawing.Point(10, 28);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(889, 20);
            this.searchTextBox.TabIndex = 3;
            // 
            // btnDeleteRecord
            // 
            this.btnDeleteRecord.Location = new System.Drawing.Point(826, 520);
            this.btnDeleteRecord.Name = "btnDeleteRecord";
            this.btnDeleteRecord.Size = new System.Drawing.Size(75, 25);
            this.btnDeleteRecord.TabIndex = 4;
            this.btnDeleteRecord.Text = "Удалить";
            this.btnDeleteRecord.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(199, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Введите ключевые слова для поиска:";
            // 
            // btnUpdateRecord
            // 
            this.btnUpdateRecord.Location = new System.Drawing.Point(726, 520);
            this.btnUpdateRecord.Name = "btnUpdateRecord";
            this.btnUpdateRecord.Size = new System.Drawing.Size(75, 25);
            this.btnUpdateRecord.TabIndex = 6;
            this.btnUpdateRecord.Text = "Обновить";
            this.btnUpdateRecord.UseVisualStyleBackColor = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 559);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(911, 22);
            this.statusStrip1.TabIndex = 7;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            // 
            // dataGridView
            // 
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(10, 56);
            this.dataGridView.Margin = new System.Windows.Forms.Padding(5);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dataGridView.Size = new System.Drawing.Size(891, 450);
            this.dataGridView.TabIndex = 2;
            // 
            // linkToTheWebsite
            // 
            this.linkToTheWebsite.AutoSize = true;
            this.linkToTheWebsite.Location = new System.Drawing.Point(112, 526);
            this.linkToTheWebsite.Name = "linkToTheWebsite";
            this.linkToTheWebsite.Size = new System.Drawing.Size(55, 13);
            this.linkToTheWebsite.TabIndex = 8;
            this.linkToTheWebsite.TabStop = true;
            this.linkToTheWebsite.Text = "linkLabel1";
            // 
            // NotPriceComponents
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(911, 581);
            this.Controls.Add(this.linkToTheWebsite);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.btnUpdateRecord);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnDeleteRecord);
            this.Controls.Add(this.searchTextBox);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.btnAddRecord);
            this.Controls.Add(this.btnWritingToSheet);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "NotPriceComponents";
            this.Text = "NotPriceComponents";
            this.TopMost = true;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.NotPriceComponents_FormClosed);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnWritingToSheet;
        private System.Windows.Forms.Button btnAddRecord;
        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.Button btnDeleteRecord;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnUpdateRecord;
        private CustomDataGridView dataGridView;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.LinkLabel linkToTheWebsite;
    }
}