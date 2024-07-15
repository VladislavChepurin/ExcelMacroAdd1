using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Iek : Vendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Iek()
        {
            RangeSearch = new[] { "IEK", "ИЕК" };
            OutValue = "Iek";
        }
    }
}
