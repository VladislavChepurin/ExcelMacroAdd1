using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class JournalNku : IJournalNku
    {
        public int Id { get; set; }
        public int Ip { get; set; }       
        public string Climate { get; set; }
        public string Reserve { get; set; }
        public string Height { get; set; }
        public string Width { get; set; }
        public string Depth { get; set; }
        public string Article { get; set; }
        public string Execution { get; set; }
        public string Vendor { get; set; }
    }
}
