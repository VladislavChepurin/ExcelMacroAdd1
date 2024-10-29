using ExcelMacroAdd.DataLayer.Entity;

namespace ExcelMacroAdd.DataLayer.Interfaces
{
    public interface ICircuitBreaker
    {
        int Id { get; set; }
        string MaxCurrent { get; set; }
        int Current { get; set; }
        string Kurve { get; set; }
        string QuantityPole { get; set; }

        // Внешний ключ
        int? ProductVendorId { get; set; }
        // Навигационное свойство
        ProductVendor ProductVendor { get; set; }

        // Внешний ключ
        int? ProductGroupId { get; set; }
        // Навигационное свойство
        ProductGroup ProductGroup { get; set; }

        // Внешний ключ
        int? ProductSeriesId { get; set; }
        // Навигационное свойство
        ProductSeries ProductSeries { get; set; }

        string ArticleNumber { get; set; }

        double WidthModule { get; set; }

        // Внешний ключ
        int? ShuntTrip24vId { get; set; }
        // Навигационное свойство
        ShuntTrip24v ShuntTrip24v { get; set; }

        // Внешний ключ
        int? ShuntTrip48vId { get; set; }
        // Навигационное свойство
        ShuntTrip48v ShuntTrip48v { get; set; }

        // Внешний ключ
        int? ShuntTrip230vId { get; set; }
        // Навигационное свойство
        ShuntTrip230v ShuntTrip230v { get; set; }

        // Внешний ключ
        int? UndervoltageReleaseId { get; set; }
        // Навигационное свойство
        UndervoltageRelease UndervoltageRelease { get; set; }

        // Внешний ключ
        int? SignalContactId { get; set; }
        // Навигационное свойство
        SignalContact SignalContact { get; set; }

        // Внешний ключ
        int? AuxiliaryContactId { get; set; }
        // Навигационное свойство
        AuxiliaryContact AuxiliaryContact { get; set; }

        // Внешний ключ
        int? SignalOrAuxiliaryContactId { get; set; }
        // Навигационное свойство
        SignalOrAuxiliaryContact SignalOrAuxiliaryContact { get; set; }
    }
}
