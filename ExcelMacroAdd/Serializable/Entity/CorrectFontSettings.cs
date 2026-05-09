using ExcelMacroAdd.Serializable.Entity.Interfaces;

namespace ExcelMacroAdd.Serializable.Entity
{
    public class CorrectFontSettings : ICorrectFontResources
    {
        public string NameFont { get; set; }
        public int SizeFont { get; set; }      
    }
}
