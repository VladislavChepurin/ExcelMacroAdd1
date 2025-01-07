using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Dkc : IVendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Dkc()
        {
            RangeSearch = new[] { "Dkc", "ДКС" };
            OutValue = "DKC";
        }
    }
}
