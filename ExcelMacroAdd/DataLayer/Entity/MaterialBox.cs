using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class MaterialBox : IMaterialBox
    {
        public int Id { get; set; }
        public string MaterialValue { get; set; }
    }
}
