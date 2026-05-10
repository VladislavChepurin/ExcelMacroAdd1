using ExcelMacroAdd.Serializable.Entity.Interfaces;

namespace ExcelMacroAdd.Serializable.Entity
{
    public class FillingOutThePassportSettings : IFillingOutThePassportSettings
    {
        public string NameFileJournal { get; set; }
        public string Template { get; set; }
        public bool CheckSHA1 { get; set; }
        public string TemplateSHA1 { get; set; }       
    }
}
