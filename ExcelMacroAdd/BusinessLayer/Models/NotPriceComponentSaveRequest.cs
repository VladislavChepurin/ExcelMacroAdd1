namespace ExcelMacroAdd.BusinessLayer.Models
{
    public sealed class NotPriceComponentSaveRequest
    {
        public string Article { get; set; }

        public string Description { get; set; }

        public string ProductVendorName { get; set; }

        public string MultiplicityName { get; set; }

        public decimal Price { get; set; }

        public int Discount { get; set; }

        public string Link { get; set; }
    }
}
