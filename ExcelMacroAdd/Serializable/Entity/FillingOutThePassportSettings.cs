using ExcelMacroAdd.Serializable.Entity.Interfaces;

namespace ExcelMacroAdd.Serializable.Entity
{
    public class FillingOutThePassportSettings : IFillingOutThePassportSettings
    {
        public string NameFileJournal { get; set; }
        public string TemplateWall { get; set; }
        public string TemplateFloor { get; set; }
        public string TemplateWallIt { get; set; }
        public string TemplateFloorIt { get; set; }


        public FillingOutThePassportSettings(string nameFileJournal, string templateWall, string templateFloor, string templateWallIt, string templateFloorIt)
        {
            NameFileJournal = nameFileJournal;
            TemplateWall = templateWall;
            TemplateFloor = templateFloor;
            TemplateWallIt = templateWallIt;
            TemplateFloorIt = templateFloorIt;
        }
    }
}
