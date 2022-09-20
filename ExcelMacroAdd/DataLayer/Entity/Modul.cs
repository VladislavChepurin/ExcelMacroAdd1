using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class Modul : IModul
    {
        public int Id { get; set; }
        public string MaxCurrent { get; set; }
        public string Current { get; set; }
        public string Kurve { get; set; }
        public string Quantity { get; set; }
        public string IekVa47 { get; set; }
        public string IekVa47m { get; set; }
        public string EkfProxima { get; set; }
        public string EkfAvers { get; set; }
        public string Keaz { get; set; }
        public string Abb { get; set; }
        public string Dkc { get; set; }
        public string Dekraft { get; set; }
        public string Schneider { get; set; }
        public string Tdm { get; set; }
        public string IekArmat { get; set; }
    }
}
