using ExcelMacroAdd.Interfaces;

namespace ExcelMacroAdd.Serializable.Entity
{
    public class Resources : IResources
    {
        public string NameFileJournal { get; set; }
        public int HeightMaxBox { get; set; }
        public string TemplateWall { get; set; }
        public string TemplateFloor { get; set; }

        public Resources(string nameFileJournal, int heightMaxBox, string templateWall, string templateFloor)
        {
            NameFileJournal = nameFileJournal;
            HeightMaxBox = heightMaxBox;
            TemplateWall = templateWall;
            TemplateFloor = templateFloor;
        }
    }
}
