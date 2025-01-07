using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Keaz : IVendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Keaz()
        {
            RangeSearch = new[] { "Keaz", "КЕАЗ" };
            OutValue = "KEAZ";
        }
    }
}
