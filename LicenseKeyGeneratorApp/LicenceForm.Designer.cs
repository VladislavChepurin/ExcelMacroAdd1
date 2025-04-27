namespace LicenseKeyGeneratorApp
{
    partial class LicenceForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            button1 = new Button();
            label1 = new Label();
            textBoxUserName = new TextBox();
            label2 = new Label();
            printDocument1 = new System.Drawing.Printing.PrintDocument();
            textBoxKey = new TextBox();
            numericUpDownYear = new NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)numericUpDownYear).BeginInit();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(110, 80);
            button1.Name = "button1";
            button1.Size = new Size(87, 23);
            button1.TabIndex = 0;
            button1.Text = "ДА!";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 9);
            label1.Name = "label1";
            label1.Size = new Size(109, 15);
            label1.TabIndex = 1;
            label1.Text = "Имя пользователя";
            // 
            // textBoxUserName
            // 
            textBoxUserName.Location = new Point(12, 27);
            textBoxUserName.Name = "textBoxUserName";
            textBoxUserName.Size = new Size(185, 23);
            textBoxUserName.TabIndex = 2;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(12, 62);
            label2.Name = "label2";
            label2.Size = new Size(117, 15);
            label2.TabIndex = 3;
            label2.Text = "Год действия ключа";
            // 
            // textBoxKey
            // 
            textBoxKey.Location = new Point(12, 109);
            textBoxKey.Name = "textBoxKey";
            textBoxKey.Size = new Size(185, 23);
            textBoxKey.TabIndex = 5;
            // 
            // numericUpDownYear
            // 
            numericUpDownYear.Location = new Point(12, 82);
            numericUpDownYear.Name = "numericUpDownYear";
            numericUpDownYear.Size = new Size(82, 23);
            numericUpDownYear.TabIndex = 6;
            // 
            // LicenceForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(210, 145);
            Controls.Add(numericUpDownYear);
            Controls.Add(textBoxKey);
            Controls.Add(label2);
            Controls.Add(textBoxUserName);
            Controls.Add(label1);
            Controls.Add(button1);
            MaximizeBox = false;
            MaximumSize = new Size(226, 184);
            MinimizeBox = false;
            MinimumSize = new Size(226, 184);
            Name = "LicenceForm";
            Text = "Генератор ключей";
            ((System.ComponentModel.ISupportInitialize)numericUpDownYear).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Label label1;
        private TextBox textBoxUserName;
        private Label label2;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private TextBox textBoxKey;
        private NumericUpDown numericUpDownYear;
    }
}
