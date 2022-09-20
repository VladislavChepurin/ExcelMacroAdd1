namespace ExcelMacroAdd.DataLayer.Interfaces
{
    public interface ISwitch
    {
        int Id { get; set; }
        string Current { get; set; }
        string Quantity { get; set; }
        string Iek { get; set; }
        string EkfProxima { get; set; }
        string EkfAvers { get; set; }
        string Keaz { get; set; }
        string Abb { get; set; }
        string Dekraft { get; set; }
        string Schneider { get; set; }
        string Tdm { get; set; }
    }
}
