namespace ExcelMacroAdd.Serializable.Entity.Interfaces
{
    public interface IFillingOutThePassportSettings
    {
        string NameFileJournal { get; set; }
        string Template { get; set; }
        bool CheckSHA1 { get; set; }
        string TemplateSHA1 { get; set; }
    }
}
