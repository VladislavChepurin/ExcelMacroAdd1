using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class Transformer: ITransformer
    {
        public int Id { get; set; }
        public string Current { get; set; }
        public string Bus { get; set; }
        public string Accuracy { get; set; }
        public string Power { get; set; }
        public string Iek { get; set; }
        public string Ekf { get; set; }
        public string Keaz { get; set; }
        public string Tdm { get; set; }
        public string IekTopTpsh { get; set; }
        public string DekraftTopTpsh { get; set; }
    }
}
