﻿using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Ekf : IVendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Ekf()
        {
            RangeSearch = new[] { "Ekf", "ЕКФ" };
            OutValue = "EKF";
        }
    }
}
