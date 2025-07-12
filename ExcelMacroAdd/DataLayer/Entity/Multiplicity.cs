using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class Multiplicity: IMultiplicity
    {
        public int Id { get; set; }
        public string Value { get; set; }      
    }
}
