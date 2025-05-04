using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Functions;
using System;
using System.ComponentModel;

namespace ExcelMacroAdd.Forms.ViewModels
{
    public class TermoCalculationViewModel: AbstractFunctions, INotifyPropertyChanged
    {
        private readonly ITermoCalcData _accessData;
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

        private const int minimumTemperature = -35;
        private const int targetTemperature = 5;

        #region Binding TextBox

        private string _textBoxHeight;
        private string _textBoxWidth;
        private string _textBoxDepth;
        private string _textBoxMinTemp;
        private string _textBoxTargetTemp;
        private string _textBoxDifferenceTemp;
        private string _textBoxHeatDissipation;

        public string TextBoxHeight
        {
            get => _textBoxHeight;
            set { _textBoxHeight = value; OnPropertyChanged(nameof(TextBoxHeight)); }
        }

        public string TextBoxWidth
        {
            get => _textBoxWidth;
            set { _textBoxWidth = value; OnPropertyChanged(nameof(TextBoxWidth)); }
        }

        public string TextBoxDepth
        {
            get => _textBoxDepth;
            set { _textBoxDepth = value; OnPropertyChanged(nameof(TextBoxDepth)); }
        }

        public string TextBoxMinTemp
        {
            get => _textBoxMinTemp;
            set { _textBoxMinTemp = value;
                OnPropertyChanged(nameof(TextBoxMinTemp));
                CalculateDifference();
            }
        }

        public string TextBoxTargetTemp
        {
            get => _textBoxTargetTemp;
            set { _textBoxTargetTemp = value; 
                OnPropertyChanged(nameof(TextBoxTargetTemp));
                CalculateDifference();
            }
        }

        public string TextBoxDifferenceTemp
        {
            get => _textBoxDifferenceTemp;
            set { _textBoxDifferenceTemp = value; OnPropertyChanged(nameof(TextBoxDifferenceTemp)); }
        }

        public string TextBoxHeatDissipation
        {
            get => _textBoxHeatDissipation;
            set { _textBoxHeatDissipation = value; OnPropertyChanged(nameof(TextBoxHeatDissipation)); }
        }

        #endregion

        #region Binding RadioButtons

        public enum Options 
        {
            ComprehensiveAccess,
            PlacementWall, 
            LastRow,
            LastRowNearWall,
            MiddleRow,
            MiddleRowNearWall,
            MiddleRowNearWallWithUpperClosed
        }

        private Options _selectedOption;

        public Options SelectedOption
        {
            get => _selectedOption;
            set
            {
                if (_selectedOption != value)
                {
                    _selectedOption = value;
                    OnPropertyChanged(nameof(SelectedOption));
                }
            }
        }

        #endregion

        #region Binding Label

        private string _textResultLabel;
        private bool _isVisibleLabel;


        public string TextResultLabel
        {
            get => _textResultLabel;
            set { _textResultLabel = value; OnPropertyChanged(nameof(TextResultLabel)); }
        }

        public bool IsVisibleLabel
        {
            get => _isVisibleLabel;
            set { _isVisibleLabel = value; OnPropertyChanged(nameof(IsVisibleLabel)); }
        }

        #endregion

        #region Binding ComboBox

        private int _installationIndex;
        private int _cabinetMaterialIndex;
        private int _insulationМaterialIndex;


        public int InstallationIndex
        {
            get => _installationIndex;
            set { _installationIndex = value; OnPropertyChanged(nameof(InstallationIndex)); }
        }

        public int CabinetMaterialIndex
        {
            get => _cabinetMaterialIndex;
            set { _cabinetMaterialIndex = value; OnPropertyChanged(nameof(CabinetMaterialIndex)); }
        }

        public int InsulationМaterialIndex
        {
            get => _insulationМaterialIndex;
            set { _insulationМaterialIndex = value; OnPropertyChanged(nameof(InsulationМaterialIndex)); }
        }

        #endregion

        public TermoCalculationViewModel(ITermoCalcData accessData)
        {
            _accessData = accessData;
            //IsVisibleLabel = true;
        }

        public async override void Start()
        {
            TextBoxMinTemp = minimumTemperature.ToString();
            TextBoxTargetTemp = targetTemperature.ToString();
            TextBoxDifferenceTemp = (targetTemperature - minimumTemperature).ToString();

            int currentRow = Cell.Row;
            string sArticle = Convert.ToString(Worksheet.Cells[currentRow, 1].Value2);

            if (sArticle == null)
                return;

            var boxData = await _accessData.AccessTermoCalc.GetEntityJournal(sArticle.ToLower());

            if (boxData == null)
                return;
                       
            TextBoxHeight = boxData.Height.ToString();
            TextBoxWidth = boxData.Width.ToString();
            TextBoxDepth = boxData.Depth.ToString();           
        }

        public void CalculateDifference()
        {
            if (int.TryParse(TextBoxMinTemp, out int lowTemp) && int.TryParse(TextBoxTargetTemp, out int internalTemp))
            {
                if (internalTemp >= lowTemp)
                {
                    TextBoxDifferenceTemp = (internalTemp - lowTemp).ToString();
                    //textBoxMinTemp.ForeColor = Color.Black;
                    //textBoxTargetTemp.ForeColor = Color.Black;
                }
                else
                {
                    TextBoxDifferenceTemp = "0";
                    //textBoxMinTemp.ForeColor = Color.Red;
                    //textBoxTargetTemp.ForeColor = Color.Red;
                }
            }
        }


        public void BthApllyClick()
        {
            double effectiveArea = GetEffectiveArea();
            double heatTransferCoefficientBox = GetHeatTransferCoefficienBox();
            double heatTransferCoefficientInsulation = GetHeatTransferCoefficienInsulation();
            double placementCoefficient = GetPlacementCoefficient();


            if (int.TryParse(TextBoxDifferenceTemp, out int temperatureDifference))
            {
                int.TryParse(TextBoxHeatDissipation, out int totalHeatGeneration);

                var heaterPower = CalculationOfHeating(placementCoefficient, effectiveArea, heatTransferCoefficientBox, heatTransferCoefficientInsulation, temperatureDifference, totalHeatGeneration);

                TextResultLabel = heaterPower.ToString() + " Вт";
                IsVisibleLabel = true;
            }
        }

        private double GetPlacementCoefficient()
        {
            if (InstallationIndex == 0)
            {
                return internalPlacement;
            }
            else if (InstallationIndex == 1)
            {
                return outdoorPlacement;
            }
            return default;
        }

        private double GetHeatTransferCoefficienBox()
        {
            if (CabinetMaterialIndex == 0)
            {
                return sheetSteel;
            }
            else if (CabinetMaterialIndex == 1)
            {
                return stainlessSteel;
            }
            else if (CabinetMaterialIndex == 2)
            {
                return seamlessPolymer;
            }
            else if (CabinetMaterialIndex == 3)
            {
                return aluminum;
            }
            return default;
        }

        private double GetHeatTransferCoefficienInsulation()
        {
            if (InsulationМaterialIndex == 0)
            {
                return GetHeatTransferCoefficienBox();
            }
            else if (InsulationМaterialIndex == 1)
            {
                return metallizedReinforcedInsulation;
            }
            else if (InsulationМaterialIndex == 2)
            {
                return doubleMetallizedReinforcedInsulation;
            }
            else if (InsulationМaterialIndex == 3)
            {
                return foamedPolyurethaneInsulation;
            }
            return default;
        }

        private double GetEffectiveArea()
        {
            if (!int.TryParse(TextBoxHeight, out int height) ||
         !int.TryParse(TextBoxWidth, out int width) ||
         !int.TryParse(TextBoxDepth, out int depth))
            {
                return 0d; // или throw new ArgumentException("Invalid dimensions");
            }

            // Преобразование в метры (разделяем логику парсинга и вычислений)
            double h = height / 1000.0;
            double w = width / 1000.0;
            double d = depth / 1000.0;

            double result;

            // Вычисление эффективной площади с помощью switch
            switch (SelectedOption)
            {
                case Options.ComprehensiveAccess:
                    result = 1.8 * h * (w + d) + 1.4 * w * d;
                    break;
                case Options.PlacementWall:
                    result = 1.4 * w * (h + d) + 1.8 * d * h;
                    break;
                case Options.LastRow:
                    result = 1.4 * d * (h + w) + 1.8 * w * h;
                    break;
                case Options.LastRowNearWall:
                    result = 1.4 * h * (w + d) + 1.4 * w * d;
                    break;
                case Options.MiddleRow:
                    result = 1.8 * w * h + 1.4 * w * d + d * h;
                    break;
                case Options.MiddleRowNearWall:
                    result = 1.4 * w * (h + d) + d * h;
                    break;
                case Options.MiddleRowNearWallWithUpperClosed:
                    result = 1.4 * w * h + 0.7 * w * d + d * h;
                    break;
                default:
                    result = 0d;
                    break;
            }
              
            return Math.Round(result, 2);
        }

        private bool TryParseDimensions(out int height, out int width, out int depth)
        {
            height = width = depth = 0;
            return int.TryParse(TextBoxHeight, out height) &&
                   int.TryParse(TextBoxWidth, out width) &&
                   int.TryParse(TextBoxDepth, out depth);
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

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
              
    }
}
