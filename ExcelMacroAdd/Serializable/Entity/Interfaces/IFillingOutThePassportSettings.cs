namespace ExcelMacroAdd.Serializable.Entity.Interfaces
{
    public interface IFillingOutThePassportSettings
    {
        string NameFileJournal { get; set; }
        int HeightMaxBox { get; set; }
        string TemplateWall { get; set; }
        string TemplateFloor { get; set; }
    }
}
