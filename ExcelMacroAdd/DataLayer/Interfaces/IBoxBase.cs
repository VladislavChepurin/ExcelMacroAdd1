using ExcelMacroAdd.DataLayer.Entity;

namespace ExcelMacroAdd.DataLayer.Interfaces
{
    public interface IBoxBase
    {
        int Id { get; set; }
        int Ip { get; set; }
        string Climate { get; set; }
        string Reserve { get; set; }
        string Height { get; set; }
        string Width { get; set; }
        string Depth { get; set; }
        string Article { get; set; }
        int ExecutionId { get; set; }
        Execution Execution { get; set; }
        int? VendorId { get; set; }
        Vendor Vendor { get; set; }
    }
}
