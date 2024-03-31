using ExcelMacroAdd.Serializable.Entity.Interfaces;

namespace ExcelMacroAdd.Serializable.Entity
{
    public class SaveState: ISaveState
    {
        public bool SaveWorkBook { get; set; }
        public bool SaveWorkSheet { get; set; }

        public SaveState(bool saveWorkBook, bool saveWorkSheet)
        {
            SaveWorkBook = saveWorkBook;
            SaveWorkSheet = saveWorkSheet;
        }
    }
}
