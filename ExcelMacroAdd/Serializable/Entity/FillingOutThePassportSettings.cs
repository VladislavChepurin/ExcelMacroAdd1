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

        public FillingOutThePassportSettings(string nameFileJournal, string templateWall, string templateFloor, string templateWallIt, string templateFloorIt,
                                             bool checkSHA1, string templateWallSHA1, string templateFloorSHA1, string templateWallItSHA1, string templateFloorItSHA1)
        {
            NameFileJournal = nameFileJournal;
            TemplateWall = templateWall;
            TemplateFloor = templateFloor;
            TemplateWallIt = templateWallIt;
            TemplateFloorIt = templateFloorIt;
            CheckSHA1 = checkSHA1;
            TemplateWallSHA1 = templateWallSHA1;
            TemplateFloorSHA1 = templateFloorSHA1;
            TemplateWallItSHA1 = templateWallItSHA1;    
            TemplateFloorItSHA1 = templateFloorItSHA1;                          
        }
    }
}
