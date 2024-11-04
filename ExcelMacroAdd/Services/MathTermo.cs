using System;

namespace ExcelMacroAdd.Services
{
    internal static class MathTermo
    {
        //Материал шкафа
        internal const double sheetSteel = 5.5;
        internal const double stainlessSteel = 4.5;
        internal const double seamlessPolymer = 3.5;
        internal const double aluminum = 12.0;
        //Материал утеплителя
        internal const double withoutInsulation = 0.0;
        internal const double metallizedReinforcedInsulation = 1.0;
        internal const double doubleMetallizedReinforcedInsulation = 0.5;
        internal const double foamedPolyurethaneInsulation = 0.2;

        //Коэффициенты размещения
        internal const double internalPlacement = 1.0;
        internal const double outdoorPlacement = 1.7;


        private static double CalculationHeatTransferCoefficient(double heatTransferCoefficientBox, double heatTransferCoefficientInsulation)
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
        internal static int CalculationOfHeating(double placementCoefficient, double effectiveArea, double heatTransferCoefficientBox, double heatTransferCoefficientInsulation, int temperatureDifference, int totalHeatGeneration)
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
        internal static double SeparatePlacement(int height, int width, int depth)
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
        internal static double LocationOnWall(int height, int width, int depth)
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
        internal static double LastPlaceInRowOfCabinets(int height, int width, int depth)
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
        internal static double LastPlaceInRowOnWall(int height, int width, int depth)
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
        internal static double LocationInMiddleOfRow(int height, int width, int depth)
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
        internal static double InMiddleOfRowOnWall(int height, int width, int depth)
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
        internal static double LocationOnWallInMiddleOfRowUnderCanopy(int height, int width, int depth)
        {
            double heightM = height / 1000.0;
            double widthM = width / 1000.0;
            double depthM = depth / 1000.0;

            var effectiveArea = 1.4 * widthM * heightM + 0.7 * widthM * depthM + depthM * heightM;
            return Math.Round(effectiveArea, 2);
        }
    }
}
