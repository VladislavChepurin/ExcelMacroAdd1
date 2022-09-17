using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.AccessLayer.Interfaces
{
    public interface IForm2Data
    {
        ISwitch GetEntitySwitch(string current, string quantity);
        IModul GetEntityModul(string current, string kurve, string maxCurrent, string quantity);
    }
}
