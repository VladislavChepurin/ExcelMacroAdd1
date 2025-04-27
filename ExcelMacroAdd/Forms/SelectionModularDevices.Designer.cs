namespace ExcelMacroAdd.Forms
{
    partial class SelectionModularDevices
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
            this.btnSelectionCircuitBreakerShow = new System.Windows.Forms.Button();
            this.btnSelectionSwitchShow = new System.Windows.Forms.Button();
            this.btnAdditionalDevicesShow = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnSelectionCircuitBreakerShow
            // 
            this.btnSelectionCircuitBreakerShow.Location = new System.Drawing.Point(21, 17);
            this.btnSelectionCircuitBreakerShow.Margin = new System.Windows.Forms.Padding(12, 8, 8, 8);
            this.btnSelectionCircuitBreakerShow.Name = "btnSelectionCircuitBreakerShow";
            this.btnSelectionCircuitBreakerShow.Size = new System.Drawing.Size(180, 50);
            this.btnSelectionCircuitBreakerShow.TabIndex = 1;
            this.btnSelectionCircuitBreakerShow.Text = "Модульные автоматические включатели";
            this.btnSelectionCircuitBreakerShow.UseVisualStyleBackColor = true;
            // 
            // btnSelectionSwitchShow
            // 
            this.btnSelectionSwitchShow.Location = new System.Drawing.Point(21, 83);
            this.btnSelectionSwitchShow.Margin = new System.Windows.Forms.Padding(8);
            this.btnSelectionSwitchShow.Name = "btnSelectionSwitchShow";
            this.btnSelectionSwitchShow.Size = new System.Drawing.Size(180, 50);
            this.btnSelectionSwitchShow.TabIndex = 2;
            this.btnSelectionSwitchShow.Text = "Модульные рубильники";
            this.btnSelectionSwitchShow.UseVisualStyleBackColor = true;
            // 
            // btnAdditionalDevicesShow
            // 
            this.btnAdditionalDevicesShow.Location = new System.Drawing.Point(21, 149);
            this.btnAdditionalDevicesShow.Margin = new System.Windows.Forms.Padding(8);
            this.btnAdditionalDevicesShow.Name = "btnAdditionalDevicesShow";
            this.btnAdditionalDevicesShow.Size = new System.Drawing.Size(180, 50);
            this.btnAdditionalDevicesShow.TabIndex = 3;
            this.btnAdditionalDevicesShow.Text = "Модульные аксесуары";
            this.btnAdditionalDevicesShow.UseVisualStyleBackColor = true;
            // 
            // SelectionModularDevices
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(222, 219);
            this.Controls.Add(this.btnAdditionalDevicesShow);
            this.Controls.Add(this.btnSelectionSwitchShow);
            this.Controls.Add(this.btnSelectionCircuitBreakerShow);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(238, 258);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(238, 258);
            this.Name = "SelectionModularDevices";
            this.Text = "Выбор модульных аппаратов";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SelectionModularDevices_FormClosed);
            this.Load += new System.EventHandler(this.SelectionModularDevices_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnSelectionCircuitBreakerShow;
        private System.Windows.Forms.Button btnSelectionSwitchShow;
        private System.Windows.Forms.Button btnAdditionalDevicesShow;
    }
}