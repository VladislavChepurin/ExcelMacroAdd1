using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Chint : Vendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Chint()
        {
            RangeSearch = new[] { "CHINT", "CH", "Чинт" };
            OutValue = "Chint";
        }
    }
}
