using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Keaz : Vendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Keaz()
        {
            RangeSearch = new[] { "KEAZ", "КЕАЗ" };
            OutValue = "Keaz";
        }
    }
}
