using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Abb : IVendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Abb()
        {
            RangeSearch = new[] { "АББ", "Abb" };
            OutValue = "ABB";
        }
    }
}
