namespace LicenseKeyGeneratorApp
{
    partial class LicenceForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.lblProduct = new System.Windows.Forms.Label();
            this.txtProduct = new System.Windows.Forms.TextBox();
            this.lblOwner = new System.Windows.Forms.Label();
            this.txtOwner = new System.Windows.Forms.TextBox();
            this.lblOrganization = new System.Windows.Forms.Label();
            this.txtOrganization = new System.Windows.Forms.TextBox();
            this.lblValidFrom = new System.Windows.Forms.Label();
            this.dtpValidFrom = new System.Windows.Forms.DateTimePicker();
            this.lblValidTo = new System.Windows.Forms.Label();
            this.dtpValidTo = new System.Windows.Forms.DateTimePicker();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.lblSignableString = new System.Windows.Forms.Label();
            this.txtSignableString = new System.Windows.Forms.TextBox();
            this.lblSignature = new System.Windows.Forms.Label();
            this.txtSignature = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // lblProduct
            // 
            this.lblProduct.AutoSize = true;
            this.lblProduct.Location = new System.Drawing.Point(12, 15);
            this.lblProduct.Name = "lblProduct";
            this.lblProduct.Size = new System.Drawing.Size(55, 15);
            this.lblProduct.Text = "Продукт:";
            // 
            // txtProduct
            // 
            this.txtProduct.Location = new System.Drawing.Point(130, 12);
            this.txtProduct.Name = "txtProduct";
            this.txtProduct.Size = new System.Drawing.Size(340, 23);
            this.txtProduct.Text = "ExcelMacroAdd";
            // 
            // lblOwner
            // 
            this.lblOwner.AutoSize = true;
            this.lblOwner.Location = new System.Drawing.Point(12, 44);
            this.lblOwner.Name = "lblOwner";
            this.lblOwner.Size = new System.Drawing.Size(63, 15);
            this.lblOwner.Text = "Владелец:";
            // 
            // txtOwner
            // 
            this.txtOwner.Location = new System.Drawing.Point(130, 41);
            this.txtOwner.Name = "txtOwner";
            this.txtOwner.Size = new System.Drawing.Size(340, 23);
            // 
            // lblOrganization
            // 
            this.lblOrganization.AutoSize = true;
            this.lblOrganization.Location = new System.Drawing.Point(12, 73);
            this.lblOrganization.Name = "lblOrganization";
            this.lblOrganization.Size = new System.Drawing.Size(82, 15);
            this.lblOrganization.Text = "Организация:";
            // 
            // txtOrganization
            // 
            this.txtOrganization.Location = new System.Drawing.Point(130, 70);
            this.txtOrganization.Name = "txtOrganization";
            this.txtOrganization.Size = new System.Drawing.Size(340, 23);
            // 
            // lblValidFrom
            // 
            this.lblValidFrom.AutoSize = true;
            this.lblValidFrom.Location = new System.Drawing.Point(12, 102);
            this.lblValidFrom.Name = "lblValidFrom";
            this.lblValidFrom.Size = new System.Drawing.Size(100, 15);
            this.lblValidFrom.Text = "Действует с:";
            // 
            // dtpValidFrom
            // 
            this.dtpValidFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpValidFrom.Location = new System.Drawing.Point(130, 99);
            this.dtpValidFrom.Name = "dtpValidFrom";
            this.dtpValidFrom.Size = new System.Drawing.Size(150, 23);
            // 
            // lblValidTo
            // 
            this.lblValidTo.AutoSize = true;
            this.lblValidTo.Location = new System.Drawing.Point(12, 131);
            this.lblValidTo.Name = "lblValidTo";
            this.lblValidTo.Size = new System.Drawing.Size(100, 15);
            this.lblValidTo.Text = "Действует до:";
            // 
            // dtpValidTo
            // 
            this.dtpValidTo.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpValidTo.Location = new System.Drawing.Point(130, 128);
            this.dtpValidTo.Name = "dtpValidTo";
            this.dtpValidTo.Size = new System.Drawing.Size(150, 23);
            this.dtpValidTo.Value = System.DateTime.Today.AddMonths(6);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnGenerate.Location = new System.Drawing.Point(130, 165);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(220, 32);
            this.btnGenerate.Text = "Сгенерировать license.json";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // lblSignableString
            // 
            this.lblSignableString.AutoSize = true;
            this.lblSignableString.Location = new System.Drawing.Point(12, 212);
            this.lblSignableString.Name = "lblSignableString";
            this.lblSignableString.Size = new System.Drawing.Size(115, 15);
            this.lblSignableString.Text = "Подписываемая строка:";
            // 
            // txtSignableString
            // 
            this.txtSignableString.Location = new System.Drawing.Point(12, 230);
            this.txtSignableString.Name = "txtSignableString";
            this.txtSignableString.ReadOnly = true;
            this.txtSignableString.Size = new System.Drawing.Size(458, 23);
            this.txtSignableString.BackColor = System.Drawing.SystemColors.Info;
            // 
            // lblSignature
            // 
            this.lblSignature.AutoSize = true;
            this.lblSignature.Location = new System.Drawing.Point(12, 262);
            this.lblSignature.Name = "lblSignature";
            this.lblSignature.Size = new System.Drawing.Size(56, 15);
            this.lblSignature.Text = "Подпись:";
            // 
            // txtSignature
            // 
            this.txtSignature.Location = new System.Drawing.Point(12, 280);
            this.txtSignature.Multiline = true;
            this.txtSignature.Name = "txtSignature";
            this.txtSignature.ReadOnly = true;
            this.txtSignature.Size = new System.Drawing.Size(458, 60);
            this.txtSignature.BackColor = System.Drawing.SystemColors.Info;
            this.txtSignature.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            // 
            // LicenceForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 355);
            this.Controls.Add(this.lblProduct);
            this.Controls.Add(this.txtProduct);
            this.Controls.Add(this.lblOwner);
            this.Controls.Add(this.txtOwner);
            this.Controls.Add(this.lblOrganization);
            this.Controls.Add(this.txtOrganization);
            this.Controls.Add(this.lblValidFrom);
            this.Controls.Add(this.dtpValidFrom);
            this.Controls.Add(this.lblValidTo);
            this.Controls.Add(this.dtpValidTo);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.lblSignableString);
            this.Controls.Add(this.txtSignableString);
            this.Controls.Add(this.lblSignature);
            this.Controls.Add(this.txtSignature);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "LicenceForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Генератор лицензий ExcelMacroAdd";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label lblProduct;
        private System.Windows.Forms.TextBox txtProduct;
        private System.Windows.Forms.Label lblOwner;
        private System.Windows.Forms.TextBox txtOwner;
        private System.Windows.Forms.Label lblOrganization;
        private System.Windows.Forms.TextBox txtOrganization;
        private System.Windows.Forms.Label lblValidFrom;
        private System.Windows.Forms.DateTimePicker dtpValidFrom;
        private System.Windows.Forms.Label lblValidTo;
        private System.Windows.Forms.DateTimePicker dtpValidTo;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Label lblSignableString;
        private System.Windows.Forms.TextBox txtSignableString;
        private System.Windows.Forms.Label lblSignature;
        private System.Windows.Forms.TextBox txtSignature;
    }
}
