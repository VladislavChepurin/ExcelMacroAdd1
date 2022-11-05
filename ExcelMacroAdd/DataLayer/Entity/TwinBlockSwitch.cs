
namespace ExcelMacroAdd.DataLayer.Entity
{
    public class TwinBlockSwitch
    {
        public int Id { get; set; }
        public string Current { get; set; }
        public bool IsReverse { get; set; }
        public string Article { get; set; }
        public byte[] Picture { get; set; }

        // Внешний ключ
        public int? DirectMountingHandleId { get; set; }
        // Навигационное свойство
        public DirectMountingHandle DirectMountingHandle { get; set; }

        // Внешний ключ
        public int? DoorHandleId { get; set; }
        // Навигационное свойство
        public DoorHandle DoorHandle { get; set; }

        // Внешний ключ
        public int? StockId { get; set; }
        // Навигационное свойство
        public Stock Stock { get; set; }

        // Внешний ключ
        public int? AdditionalPoleId { get; set; }
        // Навигационное свойство
        public AdditionalPole AdditionalPole { get; set; }
    }
}
