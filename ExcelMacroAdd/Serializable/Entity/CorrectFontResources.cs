using ExcelMacroAdd.Interfaces;

namespace ExcelMacroAdd.Serializable.Entity
{
    public class CorrectFontResources: ICorrectFontResources
    {
        public string NameFont { get; set; }
        public int SizeFont { get; set; }

        public CorrectFontResources(string nameFont, int sizeFont)
        {
            NameFont = nameFont;
            SizeFont = sizeFont;
        }
    }
}
