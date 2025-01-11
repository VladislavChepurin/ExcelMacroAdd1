﻿using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Iek : IVendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Iek()
        {
            RangeSearch = new[] { "Iek", "ИЕК" };
            OutValue = "IEK";
        }
    }
}