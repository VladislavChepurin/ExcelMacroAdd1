using ExcelMacroAdd.Interfaces;

namespace ExcelMacroAdd.Serializable.Entity
{
    public class CorrectFontSettings: ICorrectFontResources
    {
        public string NameFont { get; set; }
        public int SizeFont { get; set; }

        public CorrectFontSettings(string nameFont, int sizeFont)
        {
            NameFont = nameFont;
            SizeFont = sizeFont;
        }
    }
}
