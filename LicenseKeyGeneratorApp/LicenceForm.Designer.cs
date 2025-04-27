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
            domainUpDown1 = new DomainUpDown();
            textBoxKey = new TextBox();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(114, 80);
            button1.Name = "button1";
            button1.Size = new Size(98, 23);
            button1.TabIndex = 0;
            button1.Text = "Сгенерировать";
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
            textBoxUserName.Size = new Size(200, 23);
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
            // domainUpDown1
            // 
            domainUpDown1.Location = new Point(12, 80);
            domainUpDown1.Name = "domainUpDown1";
            domainUpDown1.Size = new Size(96, 23);
            domainUpDown1.TabIndex = 4;
            domainUpDown1.Text = "domainUpDown1";
            // 
            // textBoxKey
            // 
            textBoxKey.Location = new Point(12, 109);
            textBoxKey.Name = "textBoxKey";
            textBoxKey.Size = new Size(200, 23);
            textBoxKey.TabIndex = 5;
            // 
            // LicenceForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(220, 145);
            Controls.Add(textBoxKey);
            Controls.Add(domainUpDown1);
            Controls.Add(label2);
            Controls.Add(textBoxUserName);
            Controls.Add(label1);
            Controls.Add(button1);
            MaximizeBox = false;
            MaximumSize = new Size(236, 184);
            MinimizeBox = false;
            MinimumSize = new Size(236, 184);
            Name = "LicenceForm";
            Text = "Генератор ключей";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Label label1;
        private TextBox textBoxUserName;
        private Label label2;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private DomainUpDown domainUpDown1;
        private TextBox textBoxKey;
    }
}
