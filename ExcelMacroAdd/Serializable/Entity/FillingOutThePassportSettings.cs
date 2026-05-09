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
        public bool CheckSHA1 { get; set; }
        public string TemplateWallSHA1 { get; set; }
        public string TemplateFloorSHA1 { get; set; }
        public string TemplateWallItSHA1 { get; set; }
        public string TemplateFloorItSHA1 { get; set; }              
    }
}
