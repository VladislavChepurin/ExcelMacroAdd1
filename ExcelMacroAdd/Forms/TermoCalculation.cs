using ExcelMacroAdd.AccessLayer;
using ExcelMacroAdd.Interfaces;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class TermoCalculation : Form
    {


        private AccessData AccessJournal; 

        //Singelton
        private static TermoCalculation instance;
        public static async Task getInstance(IFormSettings formSettings, AccessData accessJournal)
        {
            if (instance == null)
            {
                await Task.Run(() =>
                {
                    instance = new TermoCalculation(accessJournal);
                    instance.TopMost = formSettings.FormTopMost;
                    instance.ShowDialog();                    
                });
            }
        }

        private TermoCalculation(AccessData accessJournal)
        {
            AccessJournal = accessJournal;
            InitializeComponent();
        }

        private void TermoCalculation_Load(object sender, EventArgs e)
        {

        }

        private void TermoCalculation_FormClosed(object sender, FormClosedEventArgs e) =>
            instance = null;


        #region KeyPress

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace        
                e.Handled = true;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace        
                e.Handled = true;
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace        
                e.Handled = true;
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8 && number != '-') // цифры и клавиша BackSpace        
                e.Handled = true;
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8 && number != '-') // цифры и клавиша BackSpace        
                e.Handled = true;
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace        
                e.Handled = true;
        }

        #endregion


    }
}
