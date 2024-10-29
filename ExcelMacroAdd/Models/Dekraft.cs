using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Dekraft : Vendors
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
