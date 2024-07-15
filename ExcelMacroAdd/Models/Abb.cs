using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Abb : Vendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Abb()
        {
            RangeSearch = new[] { "ABB", "АББ" };
            OutValue = "Abb";
        }
    }
}
