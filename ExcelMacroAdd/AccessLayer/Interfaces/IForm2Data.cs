using ExcelMacroAdd.DataLayer.Interfaces;
using System.Threading.Tasks;

namespace ExcelMacroAdd.AccessLayer.Interfaces
{
    public interface IForm2Data
    {
        Task<ISwitch> GetEntitySwitch(string current, string quantity);
        Task<IModul> GetEntityModul(string current, string kurve, string maxCurrent, string quantity);
    }
}
