using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Dekraft : IVendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Dekraft()
        {
            RangeSearch = new[] { "DEKRAFT", "DEK", "Декрафт", "Дек" };
            OutValue = "DEKraft";
        }
    }
}
