namespace ExcelMacroAdd.Forms
{
    partial class SelectionTwinBlock
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
            this.comboBoxCurrent = new System.Windows.Forms.ComboBox();
            this.checkBoxReverse = new System.Windows.Forms.CheckBox();
            this.checkBoxDirectMountingHandle = new System.Windows.Forms.CheckBox();
            this.checkBoxHandleOnDoor = new System.Windows.Forms.CheckBox();
            this.checkBoxHandleRod = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.checkBoxAdditionalPole = new System.Windows.Forms.CheckBox();
            this.btnGoSheet = new System.Windows.Forms.Button();
            this.pictureBox = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // comboBoxCurrent
            // 
            this.comboBoxCurrent.FormattingEnabled = true;
            this.comboBoxCurrent.Location = new System.Drawing.Point(278, 33);
            this.comboBoxCurrent.Name = "comboBoxCurrent";
            this.comboBoxCurrent.Size = new System.Drawing.Size(94, 21);
            this.comboBoxCurrent.TabIndex = 0;
            // 
            // checkBoxReverse
            // 
            this.checkBoxReverse.AutoSize = true;
            this.checkBoxReverse.Location = new System.Drawing.Point(378, 37);
            this.checkBoxReverse.Name = "checkBoxReverse";
            this.checkBoxReverse.Size = new System.Drawing.Size(94, 17);
            this.checkBoxReverse.TabIndex = 1;
            this.checkBoxReverse.Text = "реверсивный";
            this.checkBoxReverse.UseVisualStyleBackColor = true;
            // 
            // checkBoxDirectMountingHandle
            // 
            this.checkBoxDirectMountingHandle.AutoSize = true;
            this.checkBoxDirectMountingHandle.Location = new System.Drawing.Point(278, 97);
            this.checkBoxDirectMountingHandle.Name = "checkBoxDirectMountingHandle";
            this.checkBoxDirectMountingHandle.Size = new System.Drawing.Size(149, 17);
            this.checkBoxDirectMountingHandle.TabIndex = 2;
            this.checkBoxDirectMountingHandle.Text = "Ручка прямого монтажа";
            this.checkBoxDirectMountingHandle.UseVisualStyleBackColor = true;
            // 
            // checkBoxHandleOnDoor
            // 
            this.checkBoxHandleOnDoor.AutoSize = true;
            this.checkBoxHandleOnDoor.Location = new System.Drawing.Point(278, 120);
            this.checkBoxHandleOnDoor.Name = "checkBoxHandleOnDoor";
            this.checkBoxHandleOnDoor.Size = new System.Drawing.Size(103, 17);
            this.checkBoxHandleOnDoor.TabIndex = 3;
            this.checkBoxHandleOnDoor.Text = "Ручка на дверь";
            this.checkBoxHandleOnDoor.UseVisualStyleBackColor = true;
            // 
            // checkBoxHandleRod
            // 
            this.checkBoxHandleRod.AutoSize = true;
            this.checkBoxHandleRod.Location = new System.Drawing.Point(278, 143);
            this.checkBoxHandleRod.Name = "checkBoxHandleRod";
            this.checkBoxHandleRod.Size = new System.Drawing.Size(104, 17);
            this.checkBoxHandleRod.TabIndex = 4;
            this.checkBoxHandleRod.Text = "Шток для ручки";
            this.checkBoxHandleRod.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(275, 81);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Аксесуары";
            // 
            // checkBoxAdditionalPole
            // 
            this.checkBoxAdditionalPole.AutoSize = true;
            this.checkBoxAdditionalPole.Location = new System.Drawing.Point(278, 166);
            this.checkBoxAdditionalPole.Name = "checkBoxAdditionalPole";
            this.checkBoxAdditionalPole.Size = new System.Drawing.Size(149, 17);
            this.checkBoxAdditionalPole.TabIndex = 6;
            this.checkBoxAdditionalPole.Text = "Дополнительный полюс";
            this.checkBoxAdditionalPole.UseVisualStyleBackColor = true;
            // 
            // btnGoSheet
            // 
            this.btnGoSheet.Location = new System.Drawing.Point(278, 223);
            this.btnGoSheet.Name = "btnGoSheet";
            this.btnGoSheet.Size = new System.Drawing.Size(85, 28);
            this.btnGoSheet.TabIndex = 7;
            this.btnGoSheet.Text = "На лист";
            this.btnGoSheet.UseVisualStyleBackColor = true;
            this.btnGoSheet.Click += new System.EventHandler(this.btnGoSheet_Click);
            // 
            // pictureBox
            // 
            this.pictureBox.Location = new System.Drawing.Point(12, 12);
            this.pictureBox.Name = "pictureBox";
            this.pictureBox.Size = new System.Drawing.Size(240, 240);
            this.pictureBox.TabIndex = 8;
            this.pictureBox.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(275, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(88, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Ток рубильника";
            // 
            // SelectionTwinBlock
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(480, 263);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.pictureBox);
            this.Controls.Add(this.btnGoSheet);
            this.Controls.Add(this.checkBoxAdditionalPole);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkBoxHandleRod);
            this.Controls.Add(this.checkBoxHandleOnDoor);
            this.Controls.Add(this.checkBoxDirectMountingHandle);
            this.Controls.Add(this.checkBoxReverse);
            this.Controls.Add(this.comboBoxCurrent);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(496, 302);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(496, 302);
            this.Name = "SelectionTwinBlock";
            this.ShowIcon = false;
            this.Text = "Рубильники TwinBlock";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SelectionTwinBlock_FormClosed);
            this.Load += new System.EventHandler(this.SelectionTwinBlock_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBoxCurrent;
        private System.Windows.Forms.CheckBox checkBoxReverse;
        private System.Windows.Forms.CheckBox checkBoxDirectMountingHandle;
        private System.Windows.Forms.CheckBox checkBoxHandleOnDoor;
        private System.Windows.Forms.CheckBox checkBoxHandleRod;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBoxAdditionalPole;
        private System.Windows.Forms.Button btnGoSheet;
        private System.Windows.Forms.PictureBox pictureBox;
        private System.Windows.Forms.Label label2;
    }
}