using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class TermoCalculation : Form
    {
        private readonly ITermoCalcData accessData;
        protected readonly Worksheet Worksheet = Globals.ThisAddIn.GetActiveWorksheet();
        protected readonly Range Cell = Globals.ThisAddIn.GetActiveCell();
        static readonly Mutex Mutex = new Mutex(false, "MutexTermoCalculation_SingleInstance");
        private bool _mutexAcquired = false;

        //Материал шкафа
        private const double sheetSteel = 5.5;
        private const double stainlessSteel = 4.5;
        private const double seamlessPolymer = 3.5;
        private const double aluminum = 12.0;
        //Материал утеплителя
        private const double withoutInsulation = 0.0;
        private const double metallizedReinforcedInsulation = 1.0;
        private const double doubleMetallizedReinforcedInsulation = 0.5;
        private const double foamedPolyurethaneInsulation = 0.2;
        //Коэффициенты размещения
        private const double internalPlacement = 1.0;
        private const double outdoorPlacement = 1.7;

        public TermoCalculation(ITermoCalcData accessData, IFormSettings formSettings)
        {
            InitializeComponent();
            try
            {
                _mutexAcquired = Mutex.WaitOne(TimeSpan.FromSeconds(1), false);
                if (!_mutexAcquired)
                {
                    Close();
                }
            }
            catch (AbandonedMutexException)
            {
                _mutexAcquired = true; // Мьютекс был оставлен, но теперь принадлежит текущему потоку
            }

            TopMost = formSettings.FormTopMost;
            this.accessData = accessData;          
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

        private void TermoCalculation_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (_mutexAcquired)
            {
                Mutex.ReleaseMutex();
                _mutexAcquired = false;
            }
        }     

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

                var heaterPower = CalculationOfHeating(placementCoefficient, effectiveArea, heatTransferCoefficientBox, heatTransferCoefficientInsulation, temperatureDifference, totalHeatGeneration);

                label13.Text = heaterPower.ToString() + " Вт";
                label13.Visible = true;
            }
        }

        private double GetPlacementCoefficient()
        {
            if (comboBox1.SelectedIndex == 0)
            {
                return internalPlacement;
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                return outdoorPlacement;
            }
            return default;
        }

        private double GetHeatTransferCoefficienBox()
        {
            if (comboBox2.SelectedIndex == 0)
            {
                return sheetSteel;
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                return stainlessSteel;
            }
            else if (comboBox2.SelectedIndex == 2)
            {
                return seamlessPolymer;
            }
            else if (comboBox2.SelectedIndex == 3)
            {
                return aluminum;
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
                return metallizedReinforcedInsulation;
            }
            else if (comboBox3.SelectedIndex == 2)
            {
                return doubleMetallizedReinforcedInsulation;
            }
            else if (comboBox3.SelectedIndex == 3)
            {
                return foamedPolyurethaneInsulation;
            }
            return default;
        }

        private double GetEffectiveArea()
        {
            if (int.TryParse(textBox1.Text, out int height) && int.TryParse(textBox2.Text, out int width) && int.TryParse(textBox3.Text, out int depth))
            {
                if (radioButton1.Checked)
                {
                    return SeparatePlacement(height, width, depth);
                }
                else if (radioButton2.Checked)
                {
                    return LocationOnWall(height, width, depth);
                }
                else if (radioButton3.Checked)
                {
                    return LastPlaceInRowOfCabinets(height, width, depth);
                }
                else if (radioButton4.Checked)
                {
                    return LastPlaceInRowOnWall(height, width, depth);
                }
                else if (radioButton5.Checked)
                {
                    return LocationInMiddleOfRow(height, width, depth);
                }
                else if (radioButton6.Checked)
                {
                    return InMiddleOfRowOnWall(height, width, depth);
                }
                else if (radioButton7.Checked)
                {
                    return LocationOnWallInMiddleOfRowUnderCanopy(height, width, depth);
                }
            }
            return default;
        }

        private double CalculationHeatTransferCoefficient(double heatTransferCoefficientBox, double heatTransferCoefficientInsulation)
        {
            return (5 * heatTransferCoefficientInsulation + heatTransferCoefficientBox) / 6;
        }

        /// <summary>
        /// Расчет мощности
        /// </summary>
        /// <param name="effectiveArea"></param>
        /// <param name="heatTransferCoefficient"></param>
        /// <param name="temperatureDifference"></param>
        /// <param name="totalHeatGeneration"></param>
        /// <returns></returns>
        internal int CalculationOfHeating(double placementCoefficient, double effectiveArea, double heatTransferCoefficientBox, double heatTransferCoefficientInsulation, int temperatureDifference, int totalHeatGeneration)
        {
            var heatTransferCoefficient = CalculationHeatTransferCoefficient(heatTransferCoefficientBox, heatTransferCoefficientInsulation);

            var powerOfHeating = placementCoefficient * (effectiveArea * heatTransferCoefficient * temperatureDifference - totalHeatGeneration);
            return (int)Math.Round(powerOfHeating);
        }

        /// <summary>
        /// Отдельное размещение 
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="depth"></param>
        /// <returns></returns>
        internal double SeparatePlacement(int height, int width, int depth)
        {
            double heightM = height / 1000.0;
            double widthM = width / 1000.0;
            double depthM = depth / 1000.0;

            double effectiveArea = 1.8 * heightM * (widthM + depthM) + 1.4 * widthM * depthM;
            return Math.Round(effectiveArea, 2);
        }

        /// <summary>
        /// Расположение на стене 
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="depth"></param>
        /// <returns></returns>
        internal double LocationOnWall(int height, int width, int depth)
        {
            double heightM = height / 1000.0;
            double widthM = width / 1000.0;
            double depthM = depth / 1000.0;
            var effectiveArea = 1.4 * widthM * (heightM + depthM) + 1.8 * depthM * heightM;
            return Math.Round(effectiveArea, 2);
        }

        /// <summary>
        /// Крайнее место в ряду шкафов 
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="depth"></param>
        /// <returns></returns>
        internal double LastPlaceInRowOfCabinets(int height, int width, int depth)
        {
            double heightM = height / 1000.0;
            double widthM = width / 1000.0;
            double depthM = depth / 1000.0;

            var effectiveArea = 1.4 * depthM * (heightM + widthM) + 1.8 * widthM * heightM;
            return Math.Round(effectiveArea, 2);
        }

        /// <summary>
        /// Крайнее место в ряду на стене 
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="depth"></param>
        /// <returns></returns>
        internal double LastPlaceInRowOnWall(int height, int width, int depth)
        {
            double heightM = height / 1000.0;
            double widthM = width / 1000.0;
            double depthM = depth / 1000.0;

            var effectiveArea = 1.4 * heightM * (widthM + depthM) + 1.4 * widthM * depthM;
            return Math.Round(effectiveArea, 2);
        }

        /// <summary>
        /// Расположение в середине ряда 
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="depth"></param>
        /// <returns></returns>
        internal double LocationInMiddleOfRow(int height, int width, int depth)
        {
            double heightM = height / 1000.0;
            double widthM = width / 1000.0;
            double depthM = depth / 1000.0;

            var effectiveArea = 1.8 * widthM * heightM + 1.4 * widthM * depthM + depthM * heightM;
            return Math.Round(effectiveArea, 2);
        }

        /// <summary>
        /// В середине ряда на стене 
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="depth"></param>
        /// <returns></returns>
        internal double InMiddleOfRowOnWall(int height, int width, int depth)
        {
            double heightM = height / 1000.0;
            double widthM = width / 1000.0;
            double depthM = depth / 1000.0;

            var effectiveArea = 1.4 * widthM * (heightM + depthM) + depthM * heightM;
            return Math.Round(effectiveArea, 2);
        }

        /// <summary>
        /// Расположение на стене в середине ряда под козырьком 
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="depth"></param>
        /// <returns></returns>
        internal double LocationOnWallInMiddleOfRowUnderCanopy(int height, int width, int depth)
        {
            double heightM = height / 1000.0;
            double widthM = width / 1000.0;
            double depthM = depth / 1000.0;

            var effectiveArea = 1.4 * widthM * heightM + 0.7 * widthM * depthM + depthM * heightM;
            return Math.Round(effectiveArea, 2);
        }
    }
}
