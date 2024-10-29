using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class Switch : ISwitch
    {
        public int Id { get; set; }
        public int Current { get; set; }
        public string QuantityPole { get; set; }

        // Внешний ключ
        public int? ProductVendorId { get; set; }
        // Навигационное свойство
        public ProductVendor ProductVendor { get; set; }

        // Внешний ключ
        public int? ProductGroupId { get; set; }
        // Навигационное свойство
        public ProductGroup ProductGroup { get; set; }

        // Внешний ключ
        public int? ProductSeriesId { get; set; }
        // Навигационное свойство
        public ProductSeries ProductSeries { get; set; }

        public string ArticleNumber { get; set; }

        public double WidthModule { get; set; }

        // Внешний ключ
        public int? ShuntTrip24vId { get; set; }
         // Навигационное свойство
        public ShuntTrip24v ShuntTrip24v { get; set; }

        // Внешний ключ
        public int? ShuntTrip48vId { get; set; }
        // Навигационное свойство
        public ShuntTrip48v ShuntTrip48v { get; set; }

        // Внешний ключ
        public int? ShuntTrip230vId { get; set; }
        // Навигационное свойство
        public ShuntTrip230v ShuntTrip230v { get; set; }

        // Внешний ключ
        public int? UndervoltageReleaseId { get; set; }
        // Навигационное свойство
        public UndervoltageRelease UndervoltageRelease { get; set; }

        // Внешний ключ
        public int? SignalContactId { get; set; }
        // Навигационное свойство
        public SignalContact SignalContact { get; set; }

        // Внешний ключ
        public int? AuxiliaryContactId { get; set; }
        // Навигационное свойство
        public AuxiliaryContact AuxiliaryContact { get; set; }

        // Внешний ключ
        public int? SignalOrAuxiliaryContactId { get; set; }
        // Навигационное свойство
        public SignalOrAuxiliaryContact SignalOrAuxiliaryContact { get; set; }
    }
}
