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
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.btnDeleteRecord = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnUpdateRecord = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // btnWritingToSheet
            // 
            this.btnWritingToSheet.Location = new System.Drawing.Point(29, 530);
            this.btnWritingToSheet.Name = "btnWritingToSheet";
            this.btnWritingToSheet.Size = new System.Drawing.Size(75, 23);
            this.btnWritingToSheet.TabIndex = 0;
            this.btnWritingToSheet.Text = "На лист";
            this.btnWritingToSheet.UseVisualStyleBackColor = true;
            // 
            // btnAddRecord
            // 
            this.btnAddRecord.Location = new System.Drawing.Point(542, 530);
            this.btnAddRecord.Name = "btnAddRecord";
            this.btnAddRecord.Size = new System.Drawing.Size(75, 23);
            this.btnAddRecord.TabIndex = 1;
            this.btnAddRecord.Text = "Добавить";
            this.btnAddRecord.UseVisualStyleBackColor = true;
            // 
            // dataGridView
            // 
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(29, 54);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(742, 450);
            this.dataGridView.TabIndex = 2;
            // 
            // searchTextBox
            // 
            this.searchTextBox.Location = new System.Drawing.Point(29, 28);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(742, 20);
            this.searchTextBox.TabIndex = 3;
            // 
            // btnDeleteRecord
            // 
            this.btnDeleteRecord.Location = new System.Drawing.Point(654, 530);
            this.btnDeleteRecord.Name = "btnDeleteRecord";
            this.btnDeleteRecord.Size = new System.Drawing.Size(75, 23);
            this.btnDeleteRecord.TabIndex = 4;
            this.btnDeleteRecord.Text = "Удалить";
            this.btnDeleteRecord.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(139, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Введите ключевые слова:";
            // 
            // btnUpdateRecord
            // 
            this.btnUpdateRecord.Location = new System.Drawing.Point(438, 530);
            this.btnUpdateRecord.Name = "btnUpdateRecord";
            this.btnUpdateRecord.Size = new System.Drawing.Size(75, 23);
            this.btnUpdateRecord.TabIndex = 6;
            this.btnUpdateRecord.Text = "Обновить";
            this.btnUpdateRecord.UseVisualStyleBackColor = true;
            // 
            // NotPriceComponents
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 576);
            this.Controls.Add(this.btnUpdateRecord);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnDeleteRecord);
            this.Controls.Add(this.searchTextBox);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.btnAddRecord);
            this.Controls.Add(this.btnWritingToSheet);
            this.Name = "NotPriceComponents";
            this.Text = "NotPriceComponents";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.NotPriceComponents_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnWritingToSheet;
        private System.Windows.Forms.Button btnAddRecord;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.Button btnDeleteRecord;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnUpdateRecord;
    }
}