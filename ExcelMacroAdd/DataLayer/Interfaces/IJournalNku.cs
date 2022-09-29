namespace ExcelMacroAdd.DataLayer.Interfaces
{
    public interface IJournalNku
    {
        int Id { get; set; }
        int Ip { get; set; }
        string Climate { get; set; }
        string Reserve { get; set; }
        string Height { get; set; }
        string Width { get; set; }
        string Depth { get; set; }
        string Article { get; set; }
        string Execution { get; set; }
        string Vendor { get; set; }
    }
}
