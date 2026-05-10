using ExcelMacroAdd.DataLayer.Interfaces;
using ExcelMacroAdd.Models.Interface;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BusinessLayer.Interfaces
{
    public interface ISwitchService
    {
        Task<ISwitch> GetEntitySwitch(string vendor, string series, int current, string quantityPole);

        string[] GetAllUniqueVendors();

        string[] GetAllUniqueSeries(string vendor);

        IUserSwitch GetDataSwitch(string vendor, string series);
    }
}
