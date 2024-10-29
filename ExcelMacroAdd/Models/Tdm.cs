using ExcelMacroAdd.Models.Interface;

namespace ExcelMacroAdd.Models
{
    internal class Tdm : Vendors
    {
        public string[] RangeSearch { get; }

        public string OutValue { get; }

        public Tdm()
        {
            RangeSearch = new[] { "Tdm", "ТДМ" };
            OutValue = "TDM";
        }
    }
}
