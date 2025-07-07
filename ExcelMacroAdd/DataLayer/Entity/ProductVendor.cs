using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class ProductVendor: IProductVendor
    {
        public int Id { get; set; }
        public string VendorName { get; set; }
    }
}
