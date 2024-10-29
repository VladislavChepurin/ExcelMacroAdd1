using ExcelMacroAdd.DataLayer.Entity;

namespace ExcelMacroAdd.DataLayer.Interfaces
{
    public interface IBoxBase
    {
        int Id { get; set; }
        int Ip { get; set; }
        string Climate { get; set; }
        string Weight { get; set; }
        string Height { get; set; }
        string Width { get; set; }
        string Depth { get; set; }
        string Article { get; set; }
        int? MaterialBoxId { get; set; }
        MaterialBox MaterialBox { get; set; }
        int? ProductVendorId { get; set; }
        ProductVendor ProductVendor { get; set; }
        int? ExecutionBoxId { get; set; }
        ExecutionBox ExecutionBox { get; set; }
    }
}
