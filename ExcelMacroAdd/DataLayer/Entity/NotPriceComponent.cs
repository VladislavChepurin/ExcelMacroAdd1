using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class NotPriceComponent
    {
        private decimal? _price;

        public int Id { get; set; }
        public string Article { get; set; }
        public string Description { get; set; }

        // Внешний ключ
        public int? MultiplicityId { get; set; }
        // Навигационное свойство
        public Multiplicity Multiplicity { get; set; }

        // Внешний ключ
        public int? ProductVendorId { get; set; }
        // Навигационное свойство
        public ProductVendor ProductVendor { get; set; }

        public decimal? Price
        {
            get => _price;
            set => _price = value < 0 ? throw new ArgumentException("Цена не может быть отрицательной") : value;
        }
        public int Discount{ get; set; }

        public string DataRecord { get; set; }

        public string Link { get; set; }

        // Вычисляемое свойство для безопасного отображения вендора
        [NotMapped] // Не добавлять в базу данных
        public string VendorDisplayName => ProductVendor?.VendorName ?? "Нет вендора";

        [NotMapped] // Не добавлять в базу данных
        public string MultiplicityDisplayName => Multiplicity?.Value ?? "шт";

        [NotMapped]
        public string DataRecordDisplayName
        {
            get
            {
                if (string.IsNullOrWhiteSpace(DataRecord))
                    return "Нет даты";

                var match = Regex.Match(DataRecord, @"\b\d{2}-\d{2}-\d{4}\b");
                if (!match.Success)
                    return "Нет даты";

                // Дополнительная проверка валидности даты
                if (DateTime.TryParseExact(match.Value, "dd-MM-yyyy",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                {
                    return match.Value;
                }

                return "Неверный формат даты";
            }
        }
    }
}
