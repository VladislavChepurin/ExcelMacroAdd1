using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class NotPriceComponent
    {
        private float? _price;

        public int Id { get; set; }
        public string Article { get; set; }
        public string Description { get; set; }
        // Внешний ключ
        public int? ProductVendorId { get; set; }
        // Навигационное свойство
        public ProductVendor ProductVendor { get; set; }
        // Внешний ключ

        public double? Price
        {
            get => _price;
            set => _price = (float?)(value ?? 0f); // Преобразуем null в 0
        }
        public int Discount{ get; set; }


        // Вычисляемое свойство для безопасного отображения вендора
        [NotMapped] // Не добавлять в базу данных
        public string VendorDisplayName => ProductVendor?.VendorName ?? "Нет вендора";

    }
}
