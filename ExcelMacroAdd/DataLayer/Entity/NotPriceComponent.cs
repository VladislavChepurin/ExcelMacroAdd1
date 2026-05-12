using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class NotPriceComponent
    {
        private static readonly Regex DatePattern = new Regex(
            @"\b\d{2}-\d{2}-\d{4}\b",
            RegexOptions.Compiled);

        private decimal? _price;
        private string _cachedDataRecordDisplayName;
        private string _dataRecord;

        public int Id { get; set; }

        public int? IsValid { get; set; }

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
        public int Discount { get; set; }

        public string DataRecord
        {
            get => _dataRecord;
            set
            {
                _dataRecord = value;
                _cachedDataRecordDisplayName = null;
            }
        }

        public string Link { get; set; }

        // Вычисляемое свойство для безопасного отображения вендора
        [NotMapped] // Не добавлять в базу данных
        public string VendorDisplayName => ProductVendor?.VendorName ?? "Нет вендора";

        [NotMapped] // Не добавлять в базу данных
        public string MultiplicityDisplayName => Multiplicity?.Value ?? "шт";

        [NotMapped]
        public string DataRecordDisplayName =>
            _cachedDataRecordDisplayName ?? (_cachedDataRecordDisplayName = ParseDataRecord());

        private string ParseDataRecord()
        {
            if (string.IsNullOrWhiteSpace(_dataRecord))
                return "Нет даты";

            var match = DatePattern.Match(_dataRecord);
            if (!match.Success)
                return "Нет даты";

            if (DateTime.TryParseExact(match.Value, "dd-MM-yyyy",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            {
                return match.Value;
            }

            return "Неверный формат даты";
        }
    }
}