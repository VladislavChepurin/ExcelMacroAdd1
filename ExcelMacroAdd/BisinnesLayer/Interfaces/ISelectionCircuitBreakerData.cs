using ExcelMacroAdd.DataLayer.Interfaces;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BisinnesLayer.Interfaces
{
    public interface ISelectionCircuitBreakerData
    {
        AccessCircuitBreaker AccessCircuitBreaker { get; set; }
    }
}
