namespace ExcelMacroAdd.Serializable.Entity.Interfaces
{
    public interface IFillingOutThePassportSettings
    {
        string NameFileJournal { get; set; }
        string TemplateWall { get; set; }
        string TemplateFloor { get; set; }
        string TemplateWallIt { get; set; }
        string TemplateFloorIt { get; set; }
        bool CheckSHA1 { get; set; }
        string TemplateWallSHA1 { get; set; }
        string TemplateFloorSHA1 { get; set; }
        string TemplateWallItSHA1 { get; set; }
        string TemplateFloorItSHA1 { get; set; }
    }
}
