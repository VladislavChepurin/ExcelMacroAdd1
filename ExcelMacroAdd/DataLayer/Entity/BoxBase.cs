using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class BoxBase : IBoxBase
    {
        public int Id { get; set; }
        public int Ip { get; set; }       
        public string Climate { get; set; }
        public string Reserve { get; set; }
        public string Height { get; set; }
        public string Width { get; set; }
        public string Depth { get; set; }
        public string Article { get; set; }
        // Внешний ключ
        public int ExecutionId { get; set; }
        // Навигационное свойство
        public Execution Execution {get; set; }
        // Внешний ключ
        public int? VendorId { get; set; }
        // Навигационное свойство
        public Vendor Vendor { get; set; }
    }
}
