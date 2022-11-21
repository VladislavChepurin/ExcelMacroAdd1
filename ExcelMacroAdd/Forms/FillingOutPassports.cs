using System;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    internal partial class FillingOutPassports : Form
    {
        private readonly int countRow;

        internal FillingOutPassports(int countRow)
        {
            this.countRow = countRow;            
            InitializeComponent();
        }

        private void FillingOutPassports_Load(object sender, EventArgs e)
        {
            progressBar1.Minimum = 0;
            progressBar1.Maximum = countRow;
            progressBar1.Step = 1;
        }

        public void OnStep(int step)
        {
            this.Invoke((MethodInvoker)delegate
            {
                progressBar1.PerformStep();
                label1.Text = $@"Подождите пожайлуста, идет заполнение паспортов {step}/{countRow}.";
            });
        }

        public void OnFinal()
        {
            this.Invoke((MethodInvoker)delegate
            {
                label1.Text = @"Паспота заполнены. Ты молодец";
                button1.Enabled = true;
            });
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close(); // Закрываем форму
        }
    }
}
