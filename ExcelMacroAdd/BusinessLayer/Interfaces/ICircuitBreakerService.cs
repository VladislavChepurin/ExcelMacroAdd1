using ExcelMacroAdd.DataLayer.Interfaces;
using ExcelMacroAdd.Models.Interface;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BusinessLayer.Interfaces
{
    public interface ICircuitBreakerService
    {
        Task<ICircuitBreaker> GetEntityCircuitBreaker(string vendor, string series, int current, string curve, string maxCurrent, string quantityPole);

        string[] GetAllUniqueVendors();

        string[] GetAllUniqueSeries(string vendor);

        IUserCircuitBreaker GetDataCircuitBreaker(string vendor, string series);
    }
}
