using ExcelMacroAdd.DataLayer.Interfaces;
using System.Threading.Tasks;

namespace ExcelMacroAdd.AccessLayer.Interfaces
{
    public interface ISelectionCircuitBreakerData
    {
        AccessCircuitBreaker AccessCircuitBreaker { get; set; }
    }
}
