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
        public int? MaterialBoxId { get; set; }
        // Навигационное свойство
        public MaterialBox MaterialBox {get; set; }
        // Внешний ключ
        public int? ProductVendorId { get; set; }
        // Навигационное свойство
        public ProductVendor ProductVendor { get; set; }
        // Внешний ключ
        public int? ExecutionBoxId { get; set; }
        // Навигационное свойство
        public ExecutionBox ExecutionBox { get; set; }
    }
}
