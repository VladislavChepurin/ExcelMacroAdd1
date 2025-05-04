namespace ExcelMacroAdd.Forms
{
    partial class AdditionalDevicesForm
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
            this.checkBoxShuntTrip24V = new System.Windows.Forms.CheckBox();
            this.checkBoxShuntTrip48V = new System.Windows.Forms.CheckBox();
            this.checkBoxShuntTrip230V = new System.Windows.Forms.CheckBox();
            this.checkBoxUndervoltageRelease = new System.Windows.Forms.CheckBox();
            this.checkBoxSignalContact = new System.Windows.Forms.CheckBox();
            this.checkBoxAuxContact = new System.Windows.Forms.CheckBox();
            this.checkBoxSignalOrAuxContact = new System.Windows.Forms.CheckBox();
            this.btnApply = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // checkBoxShuntTrip24V
            // 
            this.checkBoxShuntTrip24V.AutoSize = true;
            this.checkBoxShuntTrip24V.Enabled = false;
            this.checkBoxShuntTrip24V.Location = new System.Drawing.Point(12, 12);
            this.checkBoxShuntTrip24V.Name = "checkBoxShuntTrip24V";
            this.checkBoxShuntTrip24V.Size = new System.Drawing.Size(191, 17);
            this.checkBoxShuntTrip24V.TabIndex = 0;
            this.checkBoxShuntTrip24V.Text = "Независимый расцепитель 24 В";
            this.checkBoxShuntTrip24V.UseVisualStyleBackColor = true;
            // 
            // checkBoxShuntTrip48V
            // 
            this.checkBoxShuntTrip48V.AutoSize = true;
            this.checkBoxShuntTrip48V.Enabled = false;
            this.checkBoxShuntTrip48V.Location = new System.Drawing.Point(12, 35);
            this.checkBoxShuntTrip48V.Name = "checkBoxShuntTrip48V";
            this.checkBoxShuntTrip48V.Size = new System.Drawing.Size(191, 17);
            this.checkBoxShuntTrip48V.TabIndex = 1;
            this.checkBoxShuntTrip48V.Text = "Независимый расцепитель 48 В";
            this.checkBoxShuntTrip48V.UseVisualStyleBackColor = true;
            // 
            // checkBoxShuntTrip230V
            // 
            this.checkBoxShuntTrip230V.AutoSize = true;
            this.checkBoxShuntTrip230V.Enabled = false;
            this.checkBoxShuntTrip230V.Location = new System.Drawing.Point(12, 58);
            this.checkBoxShuntTrip230V.Name = "checkBoxShuntTrip230V";
            this.checkBoxShuntTrip230V.Size = new System.Drawing.Size(197, 17);
            this.checkBoxShuntTrip230V.TabIndex = 2;
            this.checkBoxShuntTrip230V.Text = "Независимый расцепитель 230 В";
            this.checkBoxShuntTrip230V.UseVisualStyleBackColor = true;
            // 
            // checkBoxUndervoltageRelease
            // 
            this.checkBoxUndervoltageRelease.AutoSize = true;
            this.checkBoxUndervoltageRelease.Enabled = false;
            this.checkBoxUndervoltageRelease.Location = new System.Drawing.Point(12, 81);
            this.checkBoxUndervoltageRelease.Name = "checkBoxUndervoltageRelease";
            this.checkBoxUndervoltageRelease.Size = new System.Drawing.Size(235, 17);
            this.checkBoxUndervoltageRelease.TabIndex = 3;
            this.checkBoxUndervoltageRelease.Text = "Расцепитель минимального напряжения";
            this.checkBoxUndervoltageRelease.UseVisualStyleBackColor = true;
            // 
            // checkBoxSignalContact
            // 
            this.checkBoxSignalContact.AutoSize = true;
            this.checkBoxSignalContact.Enabled = false;
            this.checkBoxSignalContact.Location = new System.Drawing.Point(12, 106);
            this.checkBoxSignalContact.Name = "checkBoxSignalContact";
            this.checkBoxSignalContact.Size = new System.Drawing.Size(131, 17);
            this.checkBoxSignalContact.TabIndex = 7;
            this.checkBoxSignalContact.Text = "Сигнальный контакт";
            this.checkBoxSignalContact.UseVisualStyleBackColor = true;
            // 
            // checkBoxAuxContact
            // 
            this.checkBoxAuxContact.AutoSize = true;
            this.checkBoxAuxContact.Enabled = false;
            this.checkBoxAuxContact.Location = new System.Drawing.Point(12, 129);
            this.checkBoxAuxContact.Name = "checkBoxAuxContact";
            this.checkBoxAuxContact.Size = new System.Drawing.Size(126, 17);
            this.checkBoxAuxContact.TabIndex = 6;
            this.checkBoxAuxContact.Text = "Аварийный контакт";
            this.checkBoxAuxContact.UseVisualStyleBackColor = true;
            // 
            // checkBoxSignalOrAuxContact
            // 
            this.checkBoxSignalOrAuxContact.AutoSize = true;
            this.checkBoxSignalOrAuxContact.Enabled = false;
            this.checkBoxSignalOrAuxContact.Location = new System.Drawing.Point(12, 152);
            this.checkBoxSignalOrAuxContact.Name = "checkBoxSignalOrAuxContact";
            this.checkBoxSignalOrAuxContact.Size = new System.Drawing.Size(275, 17);
            this.checkBoxSignalOrAuxContact.TabIndex = 5;
            this.checkBoxSignalOrAuxContact.Text = "Совмещенный сигнальный и аварийный контакт";
            this.checkBoxSignalOrAuxContact.UseVisualStyleBackColor = true;
            // 
            // btnApply
            // 
            this.btnApply.Enabled = false;
            this.btnApply.Location = new System.Drawing.Point(12, 210);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(100, 25);
            this.btnApply.TabIndex = 8;
            this.btnApply.Text = "Сделать хорошо";
            this.btnApply.UseVisualStyleBackColor = true;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(182, 210);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(100, 25);
            this.btnClose.TabIndex = 9;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label1.Location = new System.Drawing.Point(5, 180);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(286, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "Дополнительные устройства на DIN-рейку не найдены";
            // 
            // AdditionalDevicesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(294, 251);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.checkBoxSignalContact);
            this.Controls.Add(this.checkBoxAuxContact);
            this.Controls.Add(this.checkBoxSignalOrAuxContact);
            this.Controls.Add(this.checkBoxUndervoltageRelease);
            this.Controls.Add(this.checkBoxShuntTrip230V);
            this.Controls.Add(this.checkBoxShuntTrip48V);
            this.Controls.Add(this.checkBoxShuntTrip24V);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(310, 290);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(310, 290);
            this.Name = "AdditionalDevicesForm";
            this.Text = "Дополнительные модульные устройства";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.AdditionalDevicesForm_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.CheckBox checkBoxShuntTrip48V;
        private System.Windows.Forms.CheckBox checkBoxShuntTrip230V;
        private System.Windows.Forms.CheckBox checkBoxUndervoltageRelease;
        private System.Windows.Forms.CheckBox checkBoxSignalContact;
        private System.Windows.Forms.CheckBox checkBoxAuxContact;
        private System.Windows.Forms.CheckBox checkBoxSignalOrAuxContact;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBoxShuntTrip24V;
    }
}