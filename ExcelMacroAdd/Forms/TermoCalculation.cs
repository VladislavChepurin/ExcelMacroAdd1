using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services;
using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class TermoCalculation : Form
    {
        private readonly ITermoCalcData accessData;
        protected readonly Worksheet Worksheet = Globals.ThisAddIn.GetActiveWorksheet();
        protected readonly Range Cell = Globals.ThisAddIn.GetActiveCell();

        //Singelton
        private static TermoCalculation instance;
        public static async Task getInstance(IFormSettings formSettings, ITermoCalcData accessData)
        {
            if (instance == null)
            {
                await Task.Run(() =>
                {
                    instance = new TermoCalculation(accessData)
                    {
                        TopMost = formSettings.FormTopMost
                    };
                    instance.ShowDialog();
                });
            }
        }

        private TermoCalculation(ITermoCalcData accessData)
        {
            this.accessData = accessData;
            InitializeComponent();
        }

        private async void TermoCalculation_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;

            int currentRow = Cell.Row;
            string sArticle = Convert.ToString(Worksheet.Cells[currentRow, 1].Value2);

            if (sArticle == null)
                return;

            var boxData = await accessData.AccessTermoCalc.GetEntityJournal(sArticle.ToLower());

            if (boxData == null)
                return;

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();

            textBox1.Text = boxData.Height.ToString();
            textBox2.Text = boxData.Width.ToString();
            textBox3.Text = boxData.Depth.ToString();
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

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (int.TryParse(textBox4.Text, out int lowTemp) && int.TryParse(textBox5.Text, out int internalTemp))
            {
                if (internalTemp >= lowTemp)
                {
                    textBox6.Text = (internalTemp - lowTemp).ToString();
                    textBox4.ForeColor = Color.Black;
                    textBox5.ForeColor = Color.Black;
                }
                else
                {
                    textBox6.Text = "0";
                    textBox4.ForeColor = Color.Red;
                    textBox5.ForeColor = Color.Red;
                }
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (int.TryParse(textBox4.Text, out int lowTemp) && int.TryParse(textBox5.Text, out int internalTemp))
            {
                if (internalTemp >= lowTemp)
                {
                    textBox6.Text = (internalTemp - lowTemp).ToString();
                    textBox4.ForeColor = Color.Black;
                    textBox5.ForeColor = Color.Black;
                }
                else
                {
                    textBox6.Text = "0";
                    textBox4.ForeColor = Color.Red;
                    textBox5.ForeColor = Color.Red;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double effectiveArea = GetEffectiveArea();
            double heatTransferCoefficientBox = GetHeatTransferCoefficienBox();
            double heatTransferCoefficientInsulation = GetHeatTransferCoefficienInsulation();
            double placementCoefficient = GetPlacementCoefficient();


            if (int.TryParse(textBox6.Text, out int temperatureDifference))
            {
                int.TryParse(textBox7.Text, out int totalHeatGeneration);

                var heaterPower = MathTermo.CalculationOfHeating(placementCoefficient, effectiveArea, heatTransferCoefficientBox, heatTransferCoefficientInsulation, temperatureDifference, totalHeatGeneration);

                label13.Text = heaterPower.ToString() + " Вт";
                label13.Visible = true;
            }
        }

        private double GetPlacementCoefficient()
        {
            if (comboBox1.SelectedIndex == 0)
            {
                return MathTermo.internalPlacement;
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                return MathTermo.outdoorPlacement;
            }
            return default;
        }


        private double GetHeatTransferCoefficienBox()
        {
            if (comboBox2.SelectedIndex == 0)
            {
                return MathTermo.sheetSteel;
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                return MathTermo.stainlessSteel;
            }
            else if (comboBox2.SelectedIndex == 2)
            {
                return MathTermo.seamlessPolymer;
            }
            else if (comboBox2.SelectedIndex == 3)
            {
                return MathTermo.aluminum;
            }
            return default;
        }

        private double GetHeatTransferCoefficienInsulation()
        {
            if (comboBox3.SelectedIndex == 0)
            {
                return GetHeatTransferCoefficienBox();
            }
            else if (comboBox3.SelectedIndex == 1)
            {
                return MathTermo.metallizedReinforcedInsulation;
            }
            else if (comboBox3.SelectedIndex == 2)
            {
                return MathTermo.doubleMetallizedReinforcedInsulation;
            }
            else if (comboBox3.SelectedIndex == 3)
            {
                return MathTermo.foamedPolyurethaneInsulation;
            }
            return default;
        }

        private double GetEffectiveArea()
        {
            if (int.TryParse(textBox1.Text, out int height) && int.TryParse(textBox2.Text, out int width) && int.TryParse(textBox3.Text, out int depth))
            {
                if (radioButton1.Checked)
                {
                    return MathTermo.SeparatePlacement(height, width, depth);
                }
                else if (radioButton2.Checked)
                {
                    return MathTermo.LocationOnWall(height, width, depth);
                }
                else if (radioButton3.Checked)
                {
                    return MathTermo.LastPlaceInRowOfCabinets(height, width, depth);
                }
                else if (radioButton4.Checked)
                {
                    return MathTermo.LastPlaceInRowOnWall(height, width, depth);
                }
                else if (radioButton5.Checked)
                {
                    return MathTermo.LocationInMiddleOfRow(height, width, depth);
                }
                else if (radioButton6.Checked)
                {
                    return MathTermo.InMiddleOfRowOnWall(height, width, depth);
                }
                else if (radioButton7.Checked)
                {
                    return MathTermo.LocationOnWallInMiddleOfRowUnderCanopy(height, width, depth);
                }
            }
            return default;
        }
    }
}
