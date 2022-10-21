namespace ExcelMacroAdd.DataLayer.Interfaces
{
    public interface ITransformer
    {
        int Id { get; set; }
        string Current { get; set; }
        string Bus { get; set; }
        string Accuracy { get; set; }
        string Power { get; set; }
        string Iek { get; set; }
        string Ekf { get; set; }
        string Keaz { get; set; }
        string Tdm { get; set; }
        string IekTopTpsh { get; set; }
        string DekraftTopTpsh { get; set; }
    }
}
