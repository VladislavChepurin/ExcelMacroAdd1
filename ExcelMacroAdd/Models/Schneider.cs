using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Schneider : IVendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Schneider()
        {
            RangeSearch = new[] { "Schneider", "Schneider Electric", "SE", "Шнайдер" };
            OutValue = "Schneider";
        }
    }
}
